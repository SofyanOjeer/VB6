VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTC 
   AutoRedraw      =   -1  'True
   Caption         =   "Trésorerie Change"
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
      TabIndex        =   46
      Top             =   0
      Width           =   3500
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   0
      TabIndex        =   18
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Menu"
      TabPicture(0)   =   "TC.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraOption"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Sélection"
      TabPicture(1)   =   "TC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgSelect"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Opération"
      TabPicture(2)   =   "TC.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraOpération"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Flux"
      TabPicture(3)   =   "TC.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picGFlux"
      Tab(3).Control(1)=   "fgFlux"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Echéancier"
      TabPicture(4)   =   "TC.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picGEch"
      Tab(4).Control(1)=   "fgEch"
      Tab(4).ControlCount=   2
      Begin VB.PictureBox picGFlux 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   700
         Left            =   -74880
         ScaleHeight     =   705
         ScaleWidth      =   9045
         TabIndex        =   56
         Top             =   5400
         Width           =   9045
      End
      Begin VB.PictureBox picGEch 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   700
         Left            =   -74880
         ScaleHeight     =   705
         ScaleWidth      =   9045
         TabIndex        =   55
         Top             =   5520
         Width           =   9045
      End
      Begin VB.Frame fraOpération 
         Height          =   5655
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   9135
         Begin VB.CommandButton cmdOpérationElpDisplay 
            Caption         =   "afficher"
            Height          =   375
            Left            =   6480
            TabIndex        =   53
            Top             =   4560
            Width           =   1215
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H00C0C0FF&
            Caption         =   "en &Attente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   5040
            Width           =   1200
         End
         Begin VB.CommandButton cmdOk 
            BackColor       =   &H00C0FFC0&
            Caption         =   "&Ok"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   7800
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   4680
            Width           =   1200
         End
         Begin VB.Frame fraOpération2 
            Caption         =   "Réglement"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3015
            Left            =   4560
            TabIndex        =   38
            Top             =   1400
            Width           =   4455
            Begin VB.CheckBox chkMontant2 
               Alignment       =   1  'Right Justify
               Caption         =   "Montant"
               Height          =   375
               Left            =   120
               TabIndex        =   57
               Top             =   480
               Width           =   900
            End
            Begin VB.CheckBox chkCorrespondantL2 
               Alignment       =   1  'Right Justify
               Caption         =   "Leur corr"
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   2500
               Width           =   1000
            End
            Begin VB.CheckBox chkCorrespondantN2 
               Alignment       =   1  'Right Justify
               Caption         =   "Notre corr"
               Height          =   255
               Left            =   120
               TabIndex        =   44
               Top             =   2000
               Width           =   1000
            End
            Begin VB.TextBox txtCorrespondantL2 
               Height          =   285
               Left            =   1200
               TabIndex        =   14
               Top             =   2520
               Width           =   2800
            End
            Begin VB.ComboBox cboDevise2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   3000
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   480
               Width           =   1400
            End
            Begin VB.TextBox txtMontant2 
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
               Height          =   285
               Left            =   1200
               TabIndex        =   11
               Top             =   480
               Width           =   1800
            End
            Begin VB.ComboBox cboCorrespondantN2 
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   2000
               Width           =   2800
            End
            Begin VB.TextBox txtIntérêts 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1200
               TabIndex        =   12
               Top             =   1000
               Width           =   1400
            End
            Begin MSComCtl2.DTPicker txtAMJEchéance 
               Height          =   300
               Left            =   1200
               TabIndex        =   9
               Top             =   1440
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
               Format          =   24707075
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label libCompte2 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               Height          =   315
               Left            =   2880
               TabIndex        =   49
               Top             =   1500
               Width           =   1395
            End
            Begin VB.Label lblAMJEchéance 
               Caption         =   "Echéance"
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   1500
               Width           =   855
            End
            Begin VB.Label lblIntérêts 
               Caption         =   "Intérêts"
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   1000
               Width           =   855
            End
         End
         Begin VB.Frame fraOpération1 
            Caption         =   "Paiement"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3015
            Left            =   120
            TabIndex        =   28
            Top             =   1400
            Width           =   4395
            Begin VB.TextBox txtCorrespondantL1 
               Height          =   285
               Left            =   960
               TabIndex        =   8
               Top             =   2520
               Width           =   2800
            End
            Begin VB.TextBox txtTaux 
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
               Height          =   285
               Left            =   960
               TabIndex        =   6
               Top             =   1000
               Width           =   1400
            End
            Begin VB.ComboBox cboCorrespondantN1 
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   2000
               Width           =   2800
            End
            Begin VB.TextBox txtMontant1 
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
               Height          =   285
               Left            =   960
               TabIndex        =   5
               Top             =   500
               Width           =   1815
            End
            Begin VB.ComboBox cboDevise1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   2880
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   500
               Width           =   1400
            End
            Begin MSComCtl2.DTPicker txtAMJValeur 
               Height          =   300
               Left            =   960
               TabIndex        =   3
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
               Format          =   24707075
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label libTaux 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               Height          =   315
               Left            =   2880
               TabIndex        =   50
               Top             =   1000
               Width           =   1400
            End
            Begin VB.Label libCompte1 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               Height          =   315
               Left            =   2880
               TabIndex        =   48
               Top             =   1500
               Width           =   1400
            End
            Begin VB.Label lblMontant1 
               Caption         =   "Montant"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   500
               Width           =   615
            End
            Begin VB.Label lblTaux 
               Caption         =   "taux"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   1005
               Width           =   855
            End
            Begin VB.Label lblCorrespondantL1 
               Caption         =   "Leur corr"
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   2500
               Width           =   735
            End
            Begin VB.Label lblCorrespondantN1 
               Caption         =   "Notre corr"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   2000
               Width           =   855
            End
            Begin VB.Label lblAMJValeur 
               Caption         =   "Valeur"
               Height          =   255
               Left            =   120
               TabIndex        =   34
               Top             =   1500
               Width           =   615
            End
         End
         Begin VB.Frame fraOpérationG 
            Caption         =   "Opération"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Width           =   8895
            Begin VB.TextBox txtRéférenceInterne 
               Height          =   285
               Left            =   1200
               TabIndex        =   54
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtEngagementCompte 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1200
               TabIndex        =   2
               Top             =   840
               Width           =   1095
            End
            Begin VB.ComboBox cboNature 
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   3240
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   360
               Width           =   2775
            End
            Begin MSComCtl2.DTPicker txtAMJEngagement 
               Height          =   300
               Left            =   7560
               TabIndex        =   1
               Top             =   360
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
               Format          =   24707075
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label libEngagementCompte 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2520
               TabIndex        =   33
               Top             =   840
               Width           =   6255
            End
            Begin VB.Label lblEngagementCompte 
               Caption         =   "Contrepartie"
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   840
               Width           =   855
            End
            Begin VB.Label lblAMJEngagement 
               Caption         =   "Date d'opération"
               Height          =   375
               Left            =   6120
               TabIndex        =   31
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label lblNature 
               Caption         =   "Nature"
               Height          =   255
               Left            =   2520
               TabIndex        =   30
               Top             =   360
               Width           =   735
            End
            Begin VB.Label lblRefInterne 
               Caption         =   "N° contrat"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   360
               Width           =   975
            End
         End
         Begin VB.Label libStatut 
            BorderStyle     =   1  'Fixed Single
            Height          =   975
            Left            =   240
            TabIndex        =   47
            Top             =   4560
            Width           =   3015
         End
         Begin VB.Label libInfo 
            BorderStyle     =   1  'Fixed Single
            Height          =   975
            Left            =   3360
            TabIndex        =   41
            Top             =   4560
            Width           =   3015
         End
      End
      Begin VB.Frame fraOption 
         Caption         =   "Options"
         Height          =   4455
         Left            =   -71280
         TabIndex        =   20
         Top             =   960
         Width           =   4575
         Begin VB.TextBox txtSelect 
            Height          =   285
            Left            =   2520
            TabIndex        =   21
            Top             =   480
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker txtAmjMax 
            Height          =   300
            Left            =   2520
            TabIndex        =   22
            Top             =   1440
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
            Format          =   24707075
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblSelect1 
            Caption         =   "Référence interne"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label lblSelect2 
            Caption         =   "Echéancier jusqu'au"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   1440
            Width           =   2295
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSelect 
         Height          =   5250
         Left            =   -74880
         TabIndex        =   19
         Top             =   720
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
         FormatString    =   $"TC.frx":008C
      End
      Begin MSFlexGridLib.MSFlexGrid fgFlux 
         Height          =   4770
         Left            =   -74880
         TabIndex        =   25
         Top             =   600
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   8414
         _Version        =   393216
         Rows            =   1
         Cols            =   11
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
         FormatString    =   $"TC.frx":01B4
      End
      Begin MSFlexGridLib.MSFlexGrid fgEch 
         Height          =   5010
         Left            =   -74880
         TabIndex        =   51
         Top             =   480
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   8837
         _Version        =   393216
         Rows            =   1
         Cols            =   11
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
         FormatString    =   $"TC.frx":02C2
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8880
      Picture         =   "TC.frx":03B5
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   15
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
      TabIndex        =   17
      Top             =   0
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuOpérationSaisir 
         Caption         =   "Saisir un contrat"
      End
      Begin VB.Menu mnuX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListàValider 
         Caption         =   "Liste des contrats à valider"
      End
      Begin VB.Menu mnuList 
         Caption         =   "Liste des contrats"
      End
      Begin VB.Menu mnuListEchéancier 
         Caption         =   "Echéancier"
      End
      Begin VB.Menu Compta 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComptaEchéancier 
         Caption         =   "Compta : échéances <= jour"
      End
      Begin VB.Menu mnuComptaLotsàValider 
         Caption         =   "Compta : lots à valider"
      End
      Begin VB.Menu mnuComptaLotComptabiliséAnnuler 
         Caption         =   "Compta : annuler un lot comptabilisé"
      End
      Begin VB.Menu X3 
         Caption         =   "-"
      End
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
         Caption         =   "Afficher ce contrat"
      End
      Begin VB.Menu mnuOpérationModifier 
         Caption         =   "Modifier ce contrat"
      End
      Begin VB.Menu mnuOpérationValider 
         Caption         =   "Valider ce contrat"
      End
      Begin VB.Menu mnuOpérationAnnuler 
         Caption         =   "Annuler ce contrat"
      End
      Begin VB.Menu mnuOpérationX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpérationEffacer 
         Caption         =   "Effacer ce contrat"
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
      Begin VB.Menu mnuEchéancierPrint 
         Caption         =   "Imprimer l'échéancier"
      End
   End
   Begin VB.Menu mnuLot 
      Caption         =   "Lot"
      Visible         =   0   'False
      Begin VB.Menu mnuLotàComptabiliserValider 
         Caption         =   "Lot à comptabiliser : valider"
      End
      Begin VB.Menu mnuLotàComptabiliserAnnuler 
         Caption         =   "Lot à comptabiliser : annuler"
      End
      Begin VB.Menu mnuLotàComptabiliserPrint 
         Caption         =   "Lot à comptabiliser : imprimer"
      End
   End
   Begin VB.Menu mnuGFlux 
      Caption         =   "Flux"
      Visible         =   0   'False
      Begin VB.Menu mnuGFluxGEch 
         Caption         =   "afficher l'échéancier"
      End
      Begin VB.Menu mnuGFluxAction 
         Caption         =   "afficher les événements"
      End
      Begin VB.Menu mnuGFluxElpDisplay 
         Caption         =   "afficher l'enregistrement"
      End
   End
   Begin VB.Menu mnuGEch 
      Caption         =   "Echéancier"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuGEchAction 
         Caption         =   "afficher l'événement"
      End
      Begin VB.Menu mnuGEchElpDisplay 
         Caption         =   "afficher l'enregistrement"
      End
   End
End
Attribute VB_Name = "frmTC"
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
Dim TCAut As typeAuthorization

Dim recTable As typeElpTable
Dim wAmjEngagement As String, wAmjEchéance As String, blnAmjEchéance As Boolean
Dim wAmjDébut  As String, wAmjFin As String
Dim paramAmjEngagementMin As String, paramAmjEngagementMax As String
Dim paramAmjEchéanceMin As String, paramAmjEchéanceMax As String
Dim wAMJEffet  As String, wAMJValeur As String

Dim fgFlux_FormatString As String, fgFlux_K As Integer
Dim fgFlux_RowDisplay As Integer, fgFlux_RowClick As Integer
Dim fgFlux_ColorClick As Long, fgFlux_ColorDisplay As Long
Dim fgFlux_Sort1 As Integer, fgFlux_Sort2 As Integer
Dim fgFlux_SortAD As Integer, fgFlux_Sort1_Old As Integer

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer

Dim fgEch_FormatString As String, fgEch_K As Integer
Dim fgEch_RowDisplay As Integer, fgEch_RowClick As Integer
Dim fgEch_ColorClick As Long, fgEch_ColorDisplay As Long
Dim fgEch_Sort1 As Integer, fgEch_Sort2 As Integer
Dim fgEch_SortAD As Integer, fgEch_Sort1_Old As Integer

Dim recGOpe As typeGOpe, xGOpe As typeGOpe, mGOpe As typeGOpe, mEchéancierGOpe As typeGOpe

Dim meGOpe() As typeGOpe
Dim meGOpe_Nb As Integer, meGOpe_Index As Integer, meGOpe_NbMax As Integer

Dim meGFlux() As typeGFlux, recGFlux As typeGFlux, mGFlux As typeGFlux
Dim meGFlux_Nb As Integer, meGFlux_Index As Integer, meGFlux_NbMax As Integer
'''Dim saveGFlux() As typeGFlux, saveGFlux_Index As Integer, saveGFlux_Nb As Integer

Dim meGECh() As typeGEch, recGEch As typeGEch, mGECh As typeGEch
Dim meGECh_Nb As Integer, meGECh_Index As Integer, meGECh_NbMax As Integer
Dim wGEch() As typeGEch, wGECh_Nb As Integer

Dim meGMemo() As typegMemo, recGMemo As typegMemo, mGMemo As typegMemo
Dim meGMemo_Nb As Integer, meGMemo_Index As Integer, meGMemo_NbMax As Integer
Dim wGMemo() As typegMemo, wGMemo_Nb As Integer, wGMemo_NbMax As Integer
Dim xGMemo() As typegMemo, xGMemo_Nb As Integer, xGMemo_NbMax As Integer

Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnfgSelect_DisplayLine As Boolean, blnfgEchéance_DisplayLine As Boolean

Dim fctPériodicité As String
Dim mAmjMin As String, mAmjMax As String, mAMJReprise As String

Dim mEChéanceCompte As String, mEngagementCompte As String, mEngagementCorrCompte As String
Dim wAmjEchéanceTrt As String * 8, wAmjExtourne As String * 8, minAmjExtourne As String * 8
Dim mNature As String, mDonneurDordre As String

Dim blnControlBiatyp As Boolean, blnComptaAuto As Boolean
Dim blnEchéancier_Gen As Boolean

Dim mOpération1_BackColor As Long, mOpération2_BackColor As Long
'$----------------------------------------------------------
Dim blnSetfocus As Boolean
Dim mCompte_Ordinaire As String * 11
Dim wCotation1_2 As String, wCotation2_1 As String


Dim paramTC As typeGParam
Public Sub fraOpération_Load(Fct As String)
Dim X As String

fgSelect_RowClick = 0
Call fgSelect_Color(fgSelect_RowDisplay, vbCyan, fgSelect_ColorClick) 'txtUsr.BackColor)
blnControl = False
mGOpe.Method = "SeekP0"
V = srvGOpe_Monitor(mGOpe)
If IsNull(V) Then
    
    lstErr.Clear: lstErr.Height = 200
    blnAmjEchéance = True
    SSTab1.Tab = 2
    ''mGOpe = xGOpe
    mGOpe.Method = Fct
    recGOpe = mGOpe
        
    GEch_Load
    If meGECh_Nb = 1 Then
        ReDim Preserve meGECh(10)
        GEch_Gen
    End If
    fgEch_Display
    
    GFlux_Load
    If meGFlux_Nb = 0 Then GFlux_Gen
    fgFlux_Display

    GMemo_Load
    
    cbo_Scan mGOpe.Nature, cboNature

    txtRéférenceInterne = mGOpe.RéférenceInterne
    
    Call DTPicker_Set(txtAmjEngagement, mGOpe.AmjEngagement): wAmjEngagement = mGOpe.AmjEngagement
    txtEngagementCompte = mId$(mGOpe.EngagementCompte, 1, 5)
    
    paramTC.NatureCode = mGOpe.Nature
    V = srvGSub_TC.paramTC_Nature(paramTC)
    '''Call srvGSub_TC.paramTC_Nature_TypeDeCompte(paramTC, mGOpe, "CptàR")
    
    cmdReset_Opération mGOpe.Nature
    cmdControl_OpérationG
    
    libTaux = cmdControl_Cotation(mGOpe.Devise1, mGOpe.Devise2, wCotation1_2, wCotation2_1)
    GSub_CV1.DeviseIso = mGOpe.Devise1
    V = CV_Attribut(GSub_CV1): If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, "? erreur devise : " & GSub_CV1.DeviseIso)
    Call srvGSub.Correspondant_cbo(cboCorrespondantN1, paramTC, GSub_CV1.DeviseN, GSub_CV1.DeviseIso, mCompte_Ordinaire)
    GSub_CV2.DeviseIso = mGOpe.Devise2
    V = CV_Attribut(GSub_CV2): If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, "? erreur devise : " & GSub_CV2.DeviseIso)
    Call srvGSub.Correspondant_cbo(cboCorrespondantN2, paramTC, GSub_CV2.DeviseN, GSub_CV2.DeviseIso, mCompte_Ordinaire)

'''''    txtMontant1 = num_Display(mGOpe.Montant1, 12, 2, Lx, X, "#")
    txtMontant1 = Format$(mGOpe.Montant1, "### ### ### ##0.00")
    cbo_Scan mGOpe.Devise1, cboDevise1
    txtTaux = Format$(mGOpe.TauxMarge1, "#####0.00000")
    txtIntérêts = Format$(mGOpe.TauxMarge2, "#####0.00000")
    Call DTPicker_Set(txtAMJValeur, mGOpe.AmjDébut): wAmjFin = mGOpe.AmjDébut
    libCompte1 = Compte_Imp(mGOpe.EngagementCompte)
    X = Compte_Imp(mGOpe.EngagementCorrCompte)
    cbo_Scan X, cboCorrespondantN1
    txtCorrespondantL1 = mGOpe.EngagementCorrSwiftL
       
    chkMontant2 = recGOpe.Flag1
    txtMontant2 = Format$(mGOpe.Montant2, "### ### ### ##0.00")
    cbo_Scan mGOpe.Devise2, cboDevise2
    Call DTPicker_Set(txtAMJEchéance, mGOpe.AmjEchéance1)
    libCompte2 = Compte_Imp(mGOpe.EchéanceCompte)
    X = Compte_Imp(mGOpe.EchéanceCorrCompte)
    cbo_Scan X, cboCorrespondantN2
    txtCorrespondantL2 = mGOpe.EchéanceCorrSwiftL
    

    currentAction = constDisplay
  
    libStatut = "Statut         : " & recStatut_Libellé(mGOpe.Statut & mGOpe.StatutPlus) & Chr$(13) _
            & "Référence  : " & Format$(mGOpe.IdRéférence, "#### ### ##0") & Chr$(13)
'                & "Saisi par     : " & meGECh(1).EchUsr & "  " & dateImp(meGECh(1).EchAMJ) & " " & timeImp(meGECh(1).EchHMS) & Chr$(13) _
'                & "Validé par   :" & meGECh(1).ActionUsr & " " & dateImp(meGECh(1).ActionAMJ) & " " & timeImp(meGECh(1).ActionHMS)

    If mGOpe.Statut = "@" Then
        Select Case mGOpe.StatutPlus
            Case Is = "V ": cmdOk.Caption = constValider
                            cmdSave.Caption = constàModifier
         
            Case Is = "? ", "  "
                        If Fct = constEffacer Then
                            cmdSave.Caption = constEffacer
                        Else
                            cmdOk.Caption = constàValider
                        End If
        End Select
    Else
    End If
''''    cmdSave.Visible = blncmdSave_Visible
    blnfgSelect_DisplayLine = True
    libRéférenceInterne = currentAction & " : " & mGOpe.Nature & " : " & mGOpe.RéférenceInterne

End If

End Sub
Public Sub GEch_Load()
Dim I As Integer

recGEch_Init recGEch

recGEch.Method = "SnapP0"
recGEch.IdRéférence = mGOpe.IdRéférence
meGECh(0) = recGEch
meGECh(0).EchSéquence = 99999
Call srvGEch_Load(recGEch, meGECh(0))

meGECh_Nb = srvGEch.arrGECh_Nb
meGECh_NbMax = meGECh_Nb + 1: ReDim meGECh(meGECh_NbMax)

For I = 1 To meGECh_Nb
    meGECh(I) = srvGEch.arrGECh(I)
    meGECh(I).Method = ""
Next I

End Sub

Public Function GEch_Save(lIdRéférence As Long)
Dim I As Integer

GEch_Save = Null

For I = 1 To meGECh_Nb
    If meGECh(I).Method = constAddNew Or meGECh(I).Method = constUpdate Then
        If meGECh(I).IdRéférence = 0 Then meGECh(I).IdRéférence = lIdRéférence
        If meGECh(I).IdRéférence = lIdRéférence Then
            V = srvGEch_Update(meGECh(I))
            If Not IsNull(V) Then GEch_Save = V
        End If
    End If
Next I

End Function


Public Function GFlux_Save(lIdRéférence As Long)
Dim I As Integer

GFlux_Save = Null

For I = 1 To meGFlux_Nb
    If meGFlux(I).Method = constAddNew Or meGFlux(I).Method = constUpdate Then
        If meGFlux(I).IdRéférence = 0 Then meGFlux(I).IdRéférence = lIdRéférence
        If meGFlux(I).IdRéférence = lIdRéférence Then
            V = srvGFlux_Update(meGFlux(I))
            If Not IsNull(V) Then GFlux_Save = V
        End If
    End If
Next I

End Function

Public Function GMemo_Save(lIdRéférence As Long)
Dim I As Integer

GMemo_Save = Null

For I = 1 To meGMemo_Nb
    If meGMemo(I).Method = constAddNew Or meGMemo(I).Method = constUpdate Then
        If meGMemo(I).IdRéférence = 0 Then meGMemo(I).IdRéférence = lIdRéférence
        If meGMemo(I).IdRéférence = lIdRéférence Then
            V = srvGMemo_Update(meGMemo(I))
            If Not IsNull(V) Then GMemo_Save = V
        End If
    End If
Next I

End Function


Public Sub GFlux_Load()
Dim I As Integer

recGFlux_Init recGFlux

recGFlux.Method = "SnapP0"
recGFlux.IdRéférence = mGOpe.IdRéférence
meGFlux(0) = recGFlux
meGFlux(0).FluxSéquence = 99999
Call srvGFlux_Load(recGFlux, meGFlux(0))

meGFlux_Nb = srvGFlux.arrGFlux_Nb
meGFlux_NbMax = meGFlux_Nb + 1: ReDim meGFlux(meGFlux_NbMax)

For I = 1 To meGFlux_Nb
    meGFlux(I) = srvGFlux.arrGFlux(I)
    meGFlux(I).Method = ""
Next I

End Sub

Public Sub GMemo_Load()
Dim I As Integer

recGMemo_Init recGMemo

recGMemo.Method = "SnapP0"
recGMemo.IdRéférence = mGOpe.IdRéférence
meGMemo(0) = recGMemo
meGMemo(0).MemoSéquence = 99999
Call srvGMemo_Load(recGMemo, meGMemo(0))

meGMemo_Nb = srvGMemo.arrgMemo_NB
meGMemo_NbMax = meGMemo_Nb + 1: ReDim meGMemo(meGMemo_NbMax)

For I = 1 To meGMemo_Nb
    meGMemo(I) = srvGMemo.arrgMemo(I)
    meGMemo(I).Method = ""
Next I

End Sub

Public Sub GFlux_Gen()

Select Case Trim(recGOpe.Nature)
    Case "CC": GFlux_GenCC
    Case "CT": GFlux_GenCT
End Select

End Sub
Public Sub GFlux_GenCC()

On Error GoTo Error_Handle

ReDim meGFlux(3): meGFlux_Nb = 2

srvGFlux.recGFlux_Init recGFlux

recGFlux.Method = constAddNew
With recGFlux                                   ' Engagement
    .IdRéférence = recGOpe.IdRéférence
    .FluxSéquence = 1
    .Application = recGOpe.Application
    .OpérationCode = "CC01"
    .Devise1 = recGOpe.Devise1
    .Montant1 = recGOpe.Montant1
    .AmjEchéanceTrt = recGOpe.AmjDébut
    .AmjDébut = recGOpe.AmjDébut
    .AmjFin = recGOpe.AmjFin
    .AmjOpération = recGOpe.AmjDébut
    .AmjValeur = recGOpe.AmjDébut
End With
meGFlux(recGFlux.FluxSéquence) = recGFlux
With recGFlux                                   ' Engagement
    .FluxSéquence = 2
    .OpérationCode = "CC51"
    .Devise1 = recGOpe.Devise2
    .Montant1 = recGOpe.Montant2
End With
meGFlux(recGFlux.FluxSéquence) = recGFlux

Exit Sub
'---------------------------------------------------------
Error_Handle:
'---------------------------------------------------------

Call MsgBox("Erreur", vbCritical, "GFlux_GenCC")


End Sub

Public Sub GFlux_GenCT()

On Error GoTo Error_Handle

ReDim meGFlux(3): meGFlux_Nb = 2

srvGFlux.recGFlux_Init recGFlux

recGFlux.Method = constAddNew
With recGFlux                                   ' Engagement
    .IdRéférence = recGOpe.IdRéférence
    .FluxSéquence = 1
    .Application = recGOpe.Application
    .OpérationCode = "CT01"
    .Devise1 = recGOpe.Devise1
    .Montant1 = recGOpe.Montant1
    .AmjEchéanceTrt = recGOpe.AmjDébut
    .AmjDébut = recGOpe.AmjDébut
    .AmjFin = recGOpe.AmjFin
    .AmjOpération = recGOpe.AmjDébut
    .AmjValeur = recGOpe.AmjDébut
End With
meGFlux(recGFlux.FluxSéquence) = recGFlux
With recGFlux                                   ' Engagement
    .FluxSéquence = 2
    .OpérationCode = "CT51"
    .Devise1 = recGOpe.Devise2
    .Montant1 = recGOpe.Montant2
End With
meGFlux(recGFlux.FluxSéquence) = recGFlux

Exit Sub
'---------------------------------------------------------
Error_Handle:
'---------------------------------------------------------

Call MsgBox("Erreur", vbCritical, "GFlux_GenCC")


End Sub

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
blnControl = False
If picGEch.Height > 700 Then: picGEch.Cls: Call pic_Resize(picGEch, 0): SSTab1.Tab = 4: Exit Sub
If picGFlux.Height > 700 Then: picGFlux.Cls: Call pic_Resize(picGFlux, 0): SSTab1.Tab = 3: Exit Sub

lstErr.Clear
If currentAction = "" Then
    If fraOption.Visible Then
        fraOption.Visible = False
    Else

        If blnMsgBox_Quit Then
        X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
        Else
           X = vbYes
        End If
        If X = vbYes Then Unload Me
    End If
Else
    If currentAction = constSaisie Then
        currentAction = constSaisieG
        Call lstErr_Clear(lstErr, lstErr, "abandon saisie ")
        fraOpérationG.Enabled = True
        fraOpération1.Enabled = False
        fraOpération2.Enabled = False
        blnControl = True
        If frmTC.Enabled Then txtEngagementCompte.SetFocus
    Else
        cmdOk.Visible = False
        cmdSave.Visible = False
        currentAction = ""
        cmdContext.Caption = constcmdRechercher
        fgSelect.Enabled = True
        fgFlux.Enabled = True
        fraOpération.Enabled = False
        If fgSelect.Rows > 1 Then
            SSTab1.Tab = 1
        Else
            cmdReset
        End If
    End If
End If

End Sub
Public Sub cmdControl()
Dim X As String, wMensualité As Currency, wAmj As String, wTaux As Double
Dim wTA As Double, wTEG As Double

If Not Me.Enabled Then Exit Sub
If SSTab1.Tab <> 2 Then Exit Sub
Me.Enabled = False

cmdOk.Visible = False
cmdSave.Visible = False
blnControl = False
blnSetfocus = False

lstErr.Clear
lstErr.Height = 200
libRéférenceInterne = currentAction & " : " & mGOpe.Nature & " : " & mGOpe.RéférenceInterne
lastActiveControl_Name = currentActiveControl_Name
xGOpe = recGOpe
recGOpe = mGOpe

recGOpe.Application = paramTC.Application
recGOpe.IPA = "E"
recGOpe.NbjBase = "0"


Select Case currentAction
    Case constSaisie:
            cmdControl_OpérationD
            GFlux_Gen
            fgFlux_Display
            GEch_Gen
            fgEch_Display
            

    Case constSaisieG
            meGECh(1).IdRéférence = 0
            blnControlBiatyp = True: cmdControl_OpérationG
    Case constDisplay
            cmdControl_OpérationD
            currentActiveControl_Name = ""
     Case constValider
            cmdControl_OpérationD
            currentActiveControl_Name = ""
           V = fctGOpe_Compare(recGOpe, mGOpe)
           If Not IsNull(V) Then Call lstErr_Clear(lstErr, cmdContext, "? " & V)
    Case Else
            blnControlBiatyp = False: Call lstErr_Clear(lstErr, cmdContext, "? action :" & currentAction)
End Select

cmdSave.Visible = blncmdSave_Visible
If lstErr.ListCount = 0 Then
    cmdOk.Visible = blncmdOk_Visible
    blnSetfocus = True: currentActiveControl_Name = "cmdOk"
End If

ExitSub:

Me.Enabled = True
    
blnControl = True

If blnSetfocus And lastActiveControl_Name <> currentActiveControl_Name Then
    Select Case currentActiveControl_Name
        Case "cmdOk": cmdOk.SetFocus
        Case "txtRéférenceInterne": txtRéférenceInterne.SetFocus
        Case "txtEngagementCompte": txtEngagementCompte.SetFocus
        Case "txtMontant1":
                            If chkMontant2 = "1" Then
                                txtMontant2.SetFocus
                            Else
                                txtMontant1.SetFocus
                           End If
        Case "txtTaux": txtTaux.SetFocus
    End Select
End If

End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
picGEch.BackColor = greenColor.BackColor
picGFlux.BackColor = greenColor.BackColor

cmdOk.Caption = constàValider: cmdOk.Visible = False
cmdSave.Caption = constEnAttente: cmdSave.Visible = False
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
blnComptaAuto = False
cmdOk.FontSize = 8: cmdOk.FontName = "MS Sans Serif"
blncmdOk_Visible = False: blncmdSave_Visible = False
blnfgSelect_DisplayLine = False: blnfgEchéance_DisplayLine = False
    
cmdOpérationElpDisplay.Visible = TCAut.Xspécial
If TCAut.Xspécial Then
    cmdSave.Height = cmdOk.Height / 2
Else
    cmdSave.Height = cmdOk.Height
End If

'fraTC.Enabled = False
fgFlux.Clear: fgFlux.Rows = 1: fgFlux_RowDisplay = 0
fgEch.Clear: fgEch.Rows = 1: fgEch_RowDisplay = 0
If cboNature.ListCount > 0 Then cboNature.ListIndex = 0
txtMontant1 = "": txtMontant2 = ""
txtTaux = ""
mAMJReprise = DSys
''wAmjEngagement = DSys: Call DTPicker_Set(txtAmjEngagement, wAmjEngagement)
''wAmjFin = DSys: Call DTPicker_Set(txtAmjEngagement, wAmjFin)
wAmjEchéance = dateFinDeMois(dateElp("MoisAdd", 1, DSys)): Call DTPicker_Set(txtAMJEchéance, wAmjEchéance)
''recRacineInit C_Racine
txtRéférenceInterne = ""
''txtRéférenceExterne = ""
txtEngagementCompte = "": libEngagementCompte = ""
''txtEchéanceCompte = "": libEchéanceCompte = ""
txtCorrespondantL1 = "": txtCorrespondantL2 = ""
libStatut = ""

''recGOpe_Init mGOpe
''mGOpe.Application = paramTC.Application
''xGOpe = mGOpe

mGOpe.Statut = "@"
mGOpe.StatutPlus = "?"
mGOpe.Method = constAddNew
mEChéanceCompte = Space$(11): mEngagementCompte = Space$(11): mEngagementCorrCompte = Space$(11)
Call DTPicker_Set(txtAmjEngagement, DSys): txtAmjEngagement.Enabled = TCAut.Xspécial
Call DTPicker_Set(txtAMJValeur, DValNext2)
Call DTPicker_Set(txtAMJEchéance, DValNext2)
'chkComptaReprise = "0"

fraOption.Visible = False
blnEchéancier_Gen = False
'''recGOpe_Init mEchéancierGOpe: mEchéancierGOpe.Application = paramTC.Application
fraOpération.Enabled = False
lastActiveControl_Name = "": currentActiveControl_Name = ""

picGEch.Cls
ReDim wGMemo(21), meGMemo(21)
wGMemo_NbMax = 20: meGMemo_NbMax = 20
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
Public Sub fgFlux_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgFlux.Row


If lRow > 0 Then
    fgFlux.Row = lRow
    For I = 0 To fgFlux.Cols - 1
        fgFlux.Col = I: fgFlux.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgFlux.Row = mRow
    If fgFlux.Row > 0 Then
        lRow = fgFlux.Row
        lColor_Old = fgFlux.CellBackColor
        For I = 0 To fgFlux.Cols - 1
          fgFlux.Col = I: fgFlux.CellBackColor = lColor
        Next I
        fgFlux.Col = 0
    End If
End If

End Sub

Public Sub fgEch_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgEch.Row


If lRow > 0 Then
    fgEch.Row = lRow
    For I = 0 To fgEch.Cols - 1
        fgEch.Col = I: fgEch.CellBackColor = lColor_Old
    Next I
    fgEch.Col = 0
End If
lRow = 0
If mRow > 0 Then
    fgEch.Row = mRow
    If fgEch.Row > 0 Then
        lRow = fgEch.Row
        lColor_Old = fgEch.CellBackColor
        For I = 0 To fgEch.Cols - 1
          fgEch.Col = I: fgEch.CellBackColor = lColor
        Next I
        fgEch.Col = 0
    End If
End If

End Sub


Public Sub fgFlux_Delete()
Call lstErr_AddItem(lstErr, cmdContext, V)
Call lstErr_AddItem(lstErr, cmdContext, "!! Suppression de l'échéancier")
meGFlux(1).Method = "DeleteAll"
'''Call srvGSub_Update(meGFlux(1))
End Sub


Public Sub fgFlux_Display()
Dim I As Integer

fgFlux.Visible = True
fgFlux.Clear: fgFlux_RowDisplay = 0: fgFlux_RowClick = 0
If picGFlux.Height > 700 Then: picGFlux.Cls: Call pic_Resize(picGFlux, 0)

fgFlux.Rows = 1
fgFlux.FormatString = fgFlux_FormatString
fgFlux.Enabled = True
For meGFlux_Index = 1 To meGFlux_Nb
    recGFlux = meGFlux(meGFlux_Index)
    fgFlux.Rows = fgFlux.Rows + 1
    fgFlux.Row = fgFlux.Rows - 1
    fgFlux_DisplayLine
Next meGFlux_Index

fgFlux_K = fgFlux.Cols
 
fgFlux_SortAD = 5
If fgFlux.Rows > 1 Then fgFlux_SortX 9
End Sub

Public Sub fgEch_Display()
Dim I As Integer

fgEch.Visible = True
fgEch.Clear: fgEch_RowDisplay = 0: fgEch_RowClick = 0
If picGEch.Height > 700 Then: picGEch.Cls: Call pic_Resize(picGEch, 0)


fgEch.Rows = 1
fgEch.FormatString = fgEch_FormatString
fgEch.Enabled = True
For meGECh_Index = 1 To meGECh_Nb
    recGEch = meGECh(meGECh_Index)
    fgEch.Rows = fgEch.Rows + 1
    fgEch.Row = fgEch.Rows - 1
    fgEch_DisplayLine
Next meGECh_Index

 
fgEch_SortAD = 5
End Sub

Public Sub fgFlux_DisplayLine()
Dim K2 As Integer

paramTC.OpérationCode = recGFlux.OpérationCode
srvGSub.param_Opération paramTC
fgFlux.Col = 0: fgFlux.Text = GSub_recOpération.Name

If recGFlux.Montant1 <> 0 Then fgFlux.Col = 1: fgFlux.Text = Format(recGFlux.Montant1, "#### ### ###.00 ") & recGFlux.Devise1
If recGFlux.Montant2 <> 0 Then fgFlux.Col = 2: fgFlux.Text = Format(recGFlux.Montant2, "#### ### ###.00 ") & recGFlux.Devise2
fgFlux.Col = 3: fgFlux.Text = dateImp(recGFlux.AmjValeur)
fgFlux.Col = 5: fgFlux.Text = "du " & dateImp(recGFlux.AmjDébut) & " au " & dateImp(recGFlux.AmjFin) & "   (" & recGFlux.Nbj & "j)"
If recGFlux.Taux <> 0 Then fgFlux.Col = 4: fgFlux.Text = Format(recGFlux.Taux, "#0.00000 ") & recGFlux.TauxProvisoire
fgFlux.Col = 6: fgFlux.Text = recStatut_Libellé(recGFlux.Statut & recGFlux.StatutPlus)
fgFlux.Col = 7: fgFlux.Text = recGFlux.IdRéférence & "_" & recGFlux.FluxSéquence
fgFlux.Col = 8: fgFlux.Text = Trim(recGFlux.Application) & "_" & Trim(recGFlux.OpérationCode)
fgFlux.Col = fgFlux.Cols - 1: fgFlux.Text = meGFlux_Index

If mId$(recGFlux.OpérationCode, 3, 2) > "50" Then
    fgFlux.Col = 1: fgFlux.CellForeColor = errUsr.ForeColor
    fgFlux.Col = 2: fgFlux.CellForeColor = errUsr.ForeColor
End If
If recGFlux.Statut = "A" Then fgFlux.Col = 0: fgFlux.CellForeColor = errUsr.ForeColor


End Sub

Public Sub fgEch_DisplayLine()
Dim wColor As Long

Select Case recGEch.Statut
    Case " ": wColor = warnUsrColor
    Case "@": wColor = greenColor.ForeColor
    Case Else: wColor = libUsr.ForeColor
End Select

fgEch.Col = 0: fgEch.Text = Trim(recGEch.Application) & "_" & Trim(recGEch.FluxSéquence): fgEch.CellForeColor = wColor
fgEch.Col = 1: fgEch.Text = recGEch.EchFct: fgEch.CellForeColor = wColor
fgEch.Col = 2: fgEch.Text = dateImp(recGEch.EchAMJ) & " - " & timeImp(recGEch.EchHMS): fgEch.CellForeColor = wColor
fgEch.Col = 3: fgEch.Text = recGEch.EchUsr: fgEch.CellForeColor = wColor
fgEch.Col = 4: fgEch.Text = recGEch.ActionFct: fgEch.CellForeColor = wColor
fgEch.Col = 5: fgEch.Text = dateImp(recGEch.ActionAmj) & " - " & timeImp(recGEch.ActionHms): fgEch.CellForeColor = wColor
fgEch.Col = 6: fgEch.Text = recGEch.ActionUsr: fgEch.CellForeColor = wColor
fgEch.Col = 7: fgEch.Text = recStatut_Libellé(recGEch.Statut & recGEch.StatutPlus): fgEch.CellForeColor = wColor
fgEch.Col = 8: fgEch.Text = recGEch.IdRéférence & "_" & recGEch.EchSéquence: fgEch.CellForeColor = wColor
fgEch.Col = fgEch.Cols - 1: fgEch.Text = meGECh_Index: fgEch.CellForeColor = wColor

If recGEch.Statut = "A" Then fgEch.Col = 0: fgEch.CellForeColor = errUsr.ForeColor

End Sub


Public Sub fgFlux_Sort()
If fgFlux.Rows > 1 Then
    fgFlux.Row = 1
    fgFlux.RowSel = fgFlux.Rows - 1
    If fgFlux_Sort1_Old = fgFlux_Sort1 Then
        If fgFlux_SortAD = 5 Then
            fgFlux_SortAD = 6
        Else
            fgFlux_SortAD = 5
        End If
    Else
        fgFlux_SortAD = 5
    End If
    fgFlux_Sort1_Old = fgFlux_Sort1
    
    fgFlux.Col = fgFlux_Sort1
    fgFlux.ColSel = fgFlux_Sort2
    fgFlux.Sort = fgFlux_SortAD
End If
    

End Sub
Public Sub fgEch_Sort()
If fgEch.Rows > 1 Then
    fgEch.Row = 1
    fgEch.RowSel = fgEch.Rows - 1
    If fgEch_Sort1_Old = fgEch_Sort1 Then
        If fgEch_SortAD = 5 Then
            fgEch_SortAD = 6
        Else
            fgEch_SortAD = 5
        End If
    Else
        fgEch_SortAD = 5
    End If
    fgEch_Sort1_Old = fgEch_Sort1
    
    fgEch.Col = fgEch_Sort1
    fgEch.ColSel = fgEch_Sort2
    fgEch.Sort = fgEch_SortAD
End If
    

End Sub

Public Sub fgFlux_SortX(lK As Integer)
Dim I As Integer
For I = 1 To fgFlux.Rows - 1
    fgFlux.Row = I
    fgFlux.Col = fgFlux.Cols - 1
    meGFlux_Index = Val(fgFlux.Text)
    fgFlux.Col = fgFlux.Cols - 2
   Select Case lK
        Case 1: fgFlux.Text = Format$(meGFlux(meGFlux_Index).Montant1, "000000000000000.00")
        Case 2: fgFlux.Text = Format$(meGFlux(meGFlux_Index).Montant2, "000000000000000.00")
        Case 3: fgFlux.Text = meGFlux(meGFlux_Index).AmjValeur
        Case 10: fgFlux.Text = Format$(meGFlux_Index, "0000000000")
    End Select
Next I

fgFlux_Sort1 = fgFlux.Cols - 2: fgFlux_Sort2 = fgFlux_Sort1
fgFlux_Sort

End Sub

Public Sub fgEch_SortX(lK As Integer)
Dim I As Integer
For I = 1 To fgEch.Rows - 1
    fgEch.Row = I
    fgEch.Col = fgEch.Cols - 1
    meGECh_Index = Val(fgEch.Text)
    fgEch.Col = fgEch.Cols - 2
    Select Case lK
        Case 2: fgEch.Text = meGECh(meGECh_Index).EchAMJ & meGECh(meGECh_Index).EchHMS
        Case 5: fgEch.Text = meGECh(meGECh_Index).ActionAmj & meGECh(meGECh_Index).ActionHms
        Case 9: fgEch.Text = Format(recGEch.IdRéférence, "000000") & "_" & Format(recGEch.EchSéquence, "0000000")
        Case 11: fgEch.Text = Format$(meGECh_Index, "0000000000")
    End Select
Next I

fgEch_Sort1 = fgEch.Cols - 2: fgEch_Sort2 = fgEch_Sort1
fgEch_Sort
End Sub

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
For meGOpe_Index = 1 To meGOpe_Nb
    If meGOpe(meGOpe_Index).Method <> constIgnore And meGOpe(meGOpe_Index).Method <> constDelete Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine
    End If
Next meGOpe_Index

fgSelect_SortAD = 5
If fgSelect.Rows = 1 Then Exit Sub
'fgSelect_Sort

End Sub
Public Sub fgSelect_DisplayLine()

fgSelect.Col = 0: fgSelect.Text = meGOpe(meGOpe_Index).RéférenceInterne
fgSelect.Col = 1: fgSelect.Text = meGOpe(meGOpe_Index).Nature
fgSelect.Col = 2: fgSelect.Text = Format(meGOpe(meGOpe_Index).Montant1, "#### ### ###.00 ")
fgSelect.Col = 3: fgSelect.Text = meGOpe(meGOpe_Index).Devise1
fgSelect.Col = 4: fgSelect.Text = dateImp(meGOpe(meGOpe_Index).AmjDébut)
fgSelect.Col = 5: fgSelect.Text = dateImp(meGOpe(meGOpe_Index).AmjFin)
fgSelect.Col = 6: fgSelect.Text = Compte_Imp(meGOpe(meGOpe_Index).EngagementCompte)
Call CV_AttributS(meGOpe(meGOpe_Index).Devise1, GSub_CV1)
recCompteInit recCompte
recCompte.Société = SocId$
recCompte.Agence = SocAgence$
recCompte.Devise = GSub_CV1.DeviseN
recCompte.Numéro = meGOpe(meGOpe_Index).EngagementCompte
mdbCptP0_Find recCompte
fgSelect.Col = 7: fgSelect.Text = recCompte.Intitulé
fgSelect.Col = 8: fgSelect.Text = meGOpe(meGOpe_Index).RéférenceExterne
fgSelect.Col = 9: fgSelect.Text = recStatut_Libellé(meGOpe(meGOpe_Index).Statut & meGOpe(meGOpe_Index).StatutPlus)
fgSelect.Col = 10: fgSelect.Text = meGOpe(meGOpe_Index).IdRéférence
fgSelect.Col = fgSelect_arrIndex - 1: fgSelect.Text = ""
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = meGOpe_Index
If meGOpe(meGOpe_Index).Statut = "@" Then
    For I = 0 To fgSelect_arrIndex
      fgSelect.Col = I: fgSelect.CellForeColor = warnUsrColor
    Next I
End If

End Sub
Public Sub fgSelect_Load()
Dim X As String, mMethod As String

recGOpe_Init xGOpe
xGOpe.Application = paramTC.Application

Select Case currentAction
    Case "mnuListàValider"
            xGOpe.Method = "SnapLA"
            xGOpe.IdRéférence = 0
            xGOpe.Statut = "@"
            
            meGOpe(0) = xGOpe
            meGOpe(0).IdRéférence = 999999999#

    Case "mnuList"
            xGOpe.Method = "SnapLRI"
            X = Trim(txtSelect)
            xGOpe.RéférenceInterne = X
            xGOpe.IdRéférence = 0
            xGOpe.Statut = " "
            
            meGOpe(0) = xGOpe
            meGOpe(0).IdRéférence = 999999999#
            meGOpe(0).RéférenceInterne = X & "9z"

End Select

Call srvGOpe_Load(xGOpe, meGOpe(0))

meGOpe_Nb = srvGOpe.arrGOpe_NB
meGOpe_NbMax = meGOpe_Nb + 1: ReDim meGOpe(meGOpe_NbMax)

For I = 1 To meGOpe_Nb
    meGOpe(I) = srvGOpe.arrGOpe(I)
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
Dim I As Integer
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    meGOpe_Index = Val(fgSelect.Text)
    fgSelect.Col = fgSelect_arrIndex - 1
    Select Case lK
        Case 1: fgSelect.Text = meGOpe(meGOpe_Index).Nature & Trim(meGOpe(meGOpe_Index).RéférenceInterne)
        Case 2: fgSelect.Text = Format$(meGOpe(meGOpe_Index).Montant1, "000000000000000.00") & meGOpe(meGOpe_Index).Devise1
        Case 3: fgSelect.Text = meGOpe(meGOpe_Index).Devise1 & Format$(meGOpe(meGOpe_Index).Montant1, "000000000000000.00")
        Case 4: fgSelect.Text = meGOpe(meGOpe_Index).AmjDébut & Trim(meGOpe(meGOpe_Index).RéférenceInterne)
        Case 5: fgSelect.Text = meGOpe(meGOpe_Index).AmjFin & Trim(meGOpe(meGOpe_Index).RéférenceInterne)
        Case fgSelect_arrIndex: fgSelect.Text = Format$(meGOpe_Index, "0000000000")
    End Select
Next I

fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub


Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents
If Not IsNull(srvGSub.param_Init(paramTC, cboNature)) Then Unload Me

Call lstErr_AddItem(lstErr, cmdContext, "Chargement des correspondants ")
DoEvents
blnRéplication_Load = True
If blnRéplication_Load Then
    srvGSub.Correspondant_LoadRéplication
Else
    srvGSub.Correspondant_LoadProduction
End If

' Chargement des devise autorisées pour les opérations de TC
recElpTable_Init recElpTable
recElpTable.Id = paramTC.TableId
recElpTable.K1 = "Devise"
Call cbo_Load(recElpTable, cboDevise1, 3)
recElpTable.Id = paramTC.TableId
recElpTable.K1 = "Devise"
Call cbo_Load(recElpTable, cboDevise2, 3)


SSTab1.Tab = 0
tableElpTable_Open
'paramAmjEngagementMin = paramAmjOpérationMin
paramAmjEngagementMax = dateElp("Ouvré", 7, DSys)
ReDim meGOpe(1)
ReDim meGECh(1): meGECh_NbMax = 1
ReDim meGFlux(1): meGFlux_NbMax = 1

cmdReset

mnuOpérationSaisir.Enabled = TCAut.Saisir
mnuListàValider.Enabled = TCAut.Consulter
mnuList.Enabled = TCAut.Consulter
mnuComptaEchéancier.Enabled = TCAut.Consulter
mnuListEchéancier.Enabled = TCAut.Consulter
mnuComptaLotsàValider.Enabled = TCAut.Comptabiliser
mnuComptaLotComptabiliséAnnuler.Enabled = TCAut.Xspécial
blnControl = False

fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = 0
fgSelect_FormatString = fgSelect.FormatString
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 12

fgFlux_Sort1 = 11: fgFlux_Sort2 = 11
fgFlux_Sort1_Old = 11
fgFlux_FormatString = fgFlux.FormatString
fgFlux_RowDisplay = 0: fgFlux_RowClick = 0

fgEch_Sort1 = 11: fgEch_Sort2 = 11
fgEch_Sort1_Old = 11
fgEch_FormatString = fgEch.FormatString
fgEch_RowDisplay = 0: fgEch_RowClick = 0


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


Private Sub cboCorrespondantN1_Click()
If blnControl Then cmdControl

End Sub


Private Sub cboCorrespondantN2_Click()
If blnControl Then cmdControl

End Sub


Private Sub cboDevise1_Click()
If blnControl Then cmdControl

End Sub

Private Sub cboDevise2_Click()
If blnControl Then cmdControl

End Sub

Private Sub cboNature_Click()
If blnControl Then cmdControl

End Sub

Private Sub cboNature_GotFocus()
lblNature.ForeColor = warnUsrColor
End Sub

Private Sub cboNature_LostFocus()
lblNature.ForeColor = lblUsr.ForeColor
'If blnControl Then cmdControl
End Sub

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdOk_Click()
Dim blnPrint As Boolean, wPrint_Msg As String
Dim V

blnPrint = False
wPrint_Msg = "Opération"

If cmdOk.Caption = constàCompta Then
'$$$    cmdSave_àCompta
Else
    cmdControl
    If lstErr.ListCount <> 0 Then Exit Sub
    Me.Enabled = False
    Select Case cmdOk.Caption
        Case constàValider
            meGECh_Nb = 1: Gech_GenUpdate
            meGFlux_Nb = 0
            recGOpe.Statut = "@"
            recGOpe.StatutPlus = "V "
            wPrint_Msg = constàValider 'cmdPrint_Call constàValider
            blnPrint = True
        Case constValider
            If Not TCAut.Xspécial And Trim(arrGECh(1).EchUsr) = Trim(usrId) Then
                Call MsgBox("Vous ne pouvez pas valider vos propres opérations.", vbCritical, "TC : Validation ")
                Call lstErr_AddItem(lstErr, cmdContext, "? validation interdite")
            Else
                cmdOk_Valider
            End If
    Case Else
            Call lstErr_AddItem(lstErr, cmdContext, "? cmdOk : " & cmdOk.Caption)
    End Select

    If lstErr.ListCount = 0 Then
        V = cmdSave_Db
    End If
    
    Me.Enabled = True
    AppActivate Me.Caption
End If

If IsNull(V) Then
    If blnPrint Then
    '$$    fraGarantie_Load " "
    '$$    cmdPrint_Call wPrint_Msg
    End If
    currentAction = "cmdOk": cmdContext_Quit
End If

End Sub

Private Sub cmdOpérationElpDisplay_Click()
srvGOpe_ElpDisplay recGOpe
End Sub

Private Sub cmdSave_Click()

cmdControl
lstErr.Clear
frmTC.Enabled = False
Select Case cmdSave.Caption
    Case constEnAttente
        meGECh_Nb = 1: Gech_GenUpdate
        meGFlux_Nb = 0
       recGOpe.Statut = "@"
        recGOpe.StatutPlus = "  "
        cmdSave_Db
        cmdContext_Quit
    Case constàModifier
        meGECh_Nb = 0
        meGFlux_Nb = 0
        recGOpe.Statut = "@"
        recGOpe.StatutPlus = "? "
        cmdSave_Db
        cmdContext_Quit
    Case constEffacer
        recGOpe.Method = constDelete
    Case Else
        Call lstErr_AddItem(lstErr, cmdContext, "? cmdsave : " & cmdSave.Caption)
End Select

''If lstErr.ListCount = 0 Then cmdSave_Db
frmTC.Enabled = True

End Sub

Private Sub fgEch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Y <= fgEch.RowHeightMin Then
    Select Case fgEch.Col
        Case 0: fgEch_Sort1 = 0: fgEch_Sort2 = 0: fgEch_Sort
        Case 1:  fgEch_Sort1 = 1: fgEch_Sort2 = 1: fgEch_Sort
        Case 2:  fgEch_SortX 2
        Case 3:  fgEch_Sort1 = 3: fgEch_Sort2 = 3: fgEch_Sort
        Case 4:  fgEch_Sort1 = 4: fgEch_Sort2 = 4: fgEch_Sort
        Case 5: fgEch_SortX 5
        Case 6: fgEch_Sort1 = 6: fgEch_Sort2 = 6: fgEch_Sort
        Case 7: fgEch_Sort1 = 7: fgEch_Sort2 = 7: fgEch_Sort
        Case 8: fgEch_Sort1 = 8: fgEch_Sort2 = 8: fgEch_Sort
        Case 10: fgEch_SortX 10
    End Select
Else
    If fgEch.Rows > 1 Then
        fgEch.Col = fgEch.Cols - 1
        meGECh_Index = Val(fgEch.Text)
        recGEch = meGECh(meGECh_Index)
        Call fgEch_Color(fgEch_RowClick, MouseMoveUsr.BackColor, fgEch_ColorClick)
        Call srvGSub.pic_Resize(picGEch, 0)
        If Button = vbRightButton Then
            mnuGEchElpDisplay.Enabled = TCAut.Xspécial
            If recGEch.FluxSéquence > 0 Then mnuGEchAction.Enabled = True
            Me.PopupMenu mnuGEch, vbPopupMenuLeftButton
        Else
            If recGEch.FluxSéquence > 0 Then mnuGEchAction_Click
        End If
          
    End If
End If

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
fgFlux.Clear: fgFlux.Row = 0
fgEch.Clear: fgEch.Row = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fgFlux_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y <= fgFlux.RowHeightMin Then
    Select Case fgFlux.Col
        Case 0: fgFlux_Sort1 = 0: fgFlux_Sort2 = 0: fgFlux_Sort
        Case 1:  fgFlux_SortX 1
        Case 2:  fgFlux_SortX 2
        Case 3:  fgFlux_SortX 3
        Case 4: fgFlux_Sort1 = 4: fgFlux_Sort2 = 4: fgFlux_Sort
        Case 6: fgFlux_Sort1 = 6: fgFlux_Sort2 = 6: fgFlux_Sort
        Case 7: fgFlux_Sort1 = 7: fgFlux_Sort2 = 7: fgFlux_Sort
        Case 8: fgFlux_Sort1 = 8: fgFlux_Sort2 = 8: fgFlux_Sort
        Case 10: fgFlux_SortX 10
    End Select
Else
    fgFlux_K = fgFlux.Row * fgFlux.Cols
    If fgFlux.Rows > 1 Then
        fgFlux.Col = fgFlux.Cols - 1
        meGFlux_Index = Val(fgFlux.Text)
        '''recGFlux.CptMvtLot = meGFlux(meGFlux_Index).CptMvtLot
        recGFlux = meGFlux(meGFlux_Index)
        Call fgFlux_Color(fgFlux_RowClick, MouseMoveUsr.BackColor, fgFlux_ColorClick)
        Call srvGSub.pic_Resize(picGFlux, 0)
       
         If Button = vbRightButton Then
            mnuGFluxElpDisplay.Enabled = TCAut.Xspécial
            mnuGFluxGEch.Enabled = True
            mnuGFluxAction.Enabled = True
           Me.PopupMenu mnuGFlux, vbPopupMenuLeftButton
        Else
            mnuGFluxAction_Click
        End If
        
        If currentAction = constDisplay Then
'$$$            Param_OpérationCode recGFlux.OpérationCode
'$$$            mnuEchéancier_Set
'$$$            Me.PopupMenu mnuEchéancier, vbPopupMenuLeftButton
           Else
'$$$            If recGFlux.CptMvtLot > 0 Then
'$$$                mnuLotàComptaValidation = False
'$$$                mnuLotàComptaAnnulation = False
'$$$                mnuLotàComptaAnnulation = False
              
 '$$$               If recGFlux.Statut = "@" And recGFlux.StatutPlus = "C " Then
 '$$$                   mnuLotàComptaValidation = TCAut.Comptabiliser
  '$$$                  mnuLotàComptaAnnulation = TCAut.Comptabiliser
  '$$$                  mnuLotàComptaPrint = TCAut.Comptabiliser
 '$$$               End If
        
  '$$$              Me.PopupMenu mnuLot, vbPopupMenuLeftButton
'$$$            End If
        End If
    End If
End If

End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xStatut As String

If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_Sort
        Case 1: fgSelect_SortX 1
        Case 2: fgSelect_SortX 2
        Case 3: fgSelect_SortX 3
        Case 4: fgSelect_SortX 4
        Case 5: fgSelect_SortX 5
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
        Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
        Case 9: fgSelect_Sort1 = 9: fgSelect_Sort2 = 9: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        fgSelect.Col = fgSelect_arrIndex
        meGOpe_Index = Val(fgSelect.Text)
        mGOpe = meGOpe(meGOpe_Index)
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    
        If mGOpe.IdRéférence > 0 Then
            mnuOpérationDisplay = TCAut.Consulter
            mnuOpérationModifier = False
            mnuOpérationAnnuler = False
            mnuOpérationEffacer = False
            mnuOpérationValider = False
          
            xStatut = mGOpe.Statut & mGOpe.StatutPlus
            If xStatut = "@? " Then
                mnuOpérationModifier = TCAut.Saisir
                mnuOpérationEffacer = TCAut.Saisir
            End If
            If xStatut = "@  " Then
                mnuOpérationModifier = TCAut.Saisir
                mnuOpérationEffacer = TCAut.Saisir
            End If
            If xStatut = "@V " Then
              If Not TCAut.Xspécial And Trim(meGECh(1).EchUsr) = Trim(usrId) Then
                    Call lstErr_Clear(lstErr, cmdContext, "! Vous ne pouvez pas valider vos opérations")
                Else
                    mnuOpérationValider = TCAut.Valider
                End If
             End If
'$$$            If xStatut = "   " Then
'$$$                mnuTCAMJFin = TCAut.Saisir
'$$$                mnuTCMainLevéePartielle = TCAut.Saisir
'$$$                mnuTCMainLevée = TCAut.Saisir
'$$$           End If
    
            Me.PopupMenu mnuOpération, vbPopupMenuLeftButton
        End If
    End If
End If

End Sub
Private Sub txtXXX_GotFocus()
'txt_GotFocus txtXXX
'End Sub
'
'Private Sub txtXXX_KeyPress(KeyAscii As Integer)
'KeyAscii = convUCase(KeyAscii)

'End Sub

'Private Sub txtXXX_LostFocus()
'txt_LostFocus txtXXX
'If blnControl Then cmdControl

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

Private Sub mnuGEchAction_Click()
If Trim(recGEch.ActionFct) <> "" Then
    V = GMemo_Scan(recGEch.EchSéquence)
    Call GMemo_Display(picGEch, wGMemo_Nb, wGMemo())
Else
    V = GFlux_Scan(recGEch.FluxSéquence)
    If IsNull(V) Then
        Call srvGSub_TC.GMemo_Gen(paramTC, recGOpe, recGFlux, recGEch, wGMemo_Nb, wGMemo())
        Call GMemo_Display(picGEch, wGMemo_Nb, wGMemo())
    End If
End If
End Sub

Private Sub mnuGEchElpDisplay_Click()
srvGEch_ElpDisplay recGEch


End Sub

Private Sub mnuGFluxAction_Click()
Dim I As Integer, J As Integer

GEch_Scan recGFlux.FluxSéquence

xGMemo_Nb = 0:
For I = 1 To wGECh_Nb
    recGEch = wGEch(I)
    If Trim(recGEch.ActionFct) <> "" Then
        V = GMemo_Scan(recGEch.EchSéquence)
    Else
        V = GFlux_Scan(recGEch.FluxSéquence)
        If IsNull(V) Then
            Call srvGSub_TC.GMemo_Gen(paramTC, recGOpe, recGFlux, recGEch, wGMemo_Nb, wGMemo())
        End If
    End If
    ReDim Preserve xGMemo(xGMemo_Nb + wGMemo_Nb + 1)
    For J = 1 To wGMemo_Nb
        xGMemo_Nb = xGMemo_Nb + 1
        xGMemo(xGMemo_Nb) = wGMemo(J)
    Next J
Next I
Call GMemo_Display(picGFlux, xGMemo_Nb, xGMemo())
End Sub

Private Sub mnuGFluxElpDisplay_Click()
srvGFlux_ElpDisplay recGFlux

End Sub

Private Sub mnuGFluxGEch_Click()
GEch_Scan recGFlux.FluxSéquence
Call srvGSub.GEch_Display(picGFlux, wGECh_Nb, wGEch())

End Sub

Private Sub mnuList_Click()
currentAction = "mnuList"
fgSelect_Load

End Sub

Private Sub mnuListàValider_Click()
currentAction = "mnuListàValider"
fgSelect_Load

End Sub

Private Sub mnuOpérationDisplay_Click()
fraOpération_Load " "
End Sub

Private Sub mnuOpérationModifier_Click()
If TCAut.Saisir Then
    mGOpe.Method = constUpdate
    mnuOpérationSaisir_Init
    fraOpération_Load "Update"
    cmdSave.Visible = True
    blncmdOk_Visible = True: blncmdSave_Visible = True
    currentAction = constSaisie
    fraOpérationG.Enabled = False
    fraOpération1.Enabled = True
    fraOpération2.Enabled = True
    blnSetfocus = True: currentActiveControl_Name = "txtMontant1"
    blnControl = True: cmdControl
End If

End Sub

Private Sub mnuOpérationSaisir_Click()
If TCAut.Saisir Then
    recGOpe_Init mGOpe
    mGOpe.Method = constAddNew
    currentAction = constSaisieG
    mnuOpérationSaisir_Init
End If

End Sub


Private Sub mnuOpérationValider_Click()
If TCAut.Valider Then
    mGOpe.Method = constUpdate
    mnuOpérationSaisir_Init
    fraOpération_Load "Update"
    blncmdOk_Visible = True: blncmdSave_Visible = True
    cmdOk.Visible = True
    cmdSave.Visible = True
    currentAction = constValider
    cmdContext.Caption = constcmdAbandonner
    fraOpérationG.Enabled = False
    fraOpération.Enabled = True
End If
End Sub

Private Sub chkmontant2_Click()
If blnControl Then cmdControl

End Sub

Private Sub txtEngagementCompte_GotFocus()
txt_GotFocus txtEngagementCompte
End Sub

Private Sub txtEngagementCompte_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub

Private Sub txtEngagementCompte_LostFocus()
txt_LostFocus txtEngagementCompte
If blnControl Then cmdControl

End Sub
Private Sub txtmontant1_GotFocus()
txt_GotFocus txtMontant1
End Sub

Private Sub txtmontant1_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtMontant1)
End Sub
Private Sub txtmontant1_LostFocus()
txt_LostFocus txtMontant1
If blnControl Then cmdControl

End Sub

Private Sub txtMontant2_GotFocus()
txt_GotFocus txtMontant2
End Sub

Private Sub txtMontant2_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtMontant2)
End Sub

Private Sub txtMontant2_LostFocus()
txt_LostFocus txtMontant2
If blnControl Then cmdControl

End Sub

Private Sub txtTaux_GotFocus()
txt_GotFocus txtTaux
End Sub

Private Sub txtTaux_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtTaux)

End Sub

Private Sub txtTaux_LostFocus()
txt_LostFocus txtTaux
If blnControl Then cmdControl

End Sub
Private Sub txtIntérêts_GotFocus()
txt_GotFocus txtIntérêts
End Sub

Private Sub txtIntérêts_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtIntérêts_LostFocus()
txt_LostFocus txtIntérêts
If blnControl Then cmdControl

End Sub
Private Sub txtCorrespondantL1_GotFocus()
txt_GotFocus txtCorrespondantL1
End Sub

Private Sub txtCorrespondantL1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtCorrespondantL1_LostFocus()
txt_LostFocus txtCorrespondantL1
If blnControl Then cmdControl

End Sub
Private Sub txtCorrespondantL2_GotFocus()
txt_GotFocus txtCorrespondantL2
End Sub

Private Sub txtCorrespondantL2_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtCorrespondantL2_LostFocus()
txt_LostFocus txtCorrespondantL2
If blnControl Then cmdControl

End Sub






Private Sub txtRéférenceInterne_GotFocus()
txt_GotFocus txtRéférenceInterne

End Sub


Private Sub txtRéférenceInterne_LostFocus()
txt_LostFocus txtRéférenceInterne
If blnControl Then cmdControl

End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

paramTC.TableId = "GFlux_TC"
Call BiaPgmAut_Init("TC", TCAut)
mnuOpérationSaisir.Enabled = TCAut.Saisir

Form_Init
'If UCase$(Trim(mId$(Msg, 13, 12))) = "BIA_EXPLOIT" Then
'    mnuListEchéancier_Click
'    mnucmdPrintList_TFlux_Click
'    Unload Me
'End If

End Sub


Public Sub cmdControl_OpérationG()

X = Trim(txtRéférenceInterne)
If X = "" Then
    blnSetfocus = True: currentActiveControl_Name = "txtRéférenceInterne"
    Call lstErr_AddItem(lstErr, lstErr, "? préciser le N° du contrat")
Else
    recGOpe.RéférenceInterne = X
End If

If Not IsNull(srvGSub.param_Init(paramTC, cboNature)) Then Unload Me

If recGOpe.Nature <> xGOpe.Nature Then
   ''' txtEngagementCompte = ""
    cmdReset_Opération recGOpe.Nature
End If
paramTC.NatureCode = recGOpe.Nature
V = srvGSub_TC.paramTC_Nature(paramTC)
If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, V)

Call DTPicker_Control(txtAmjEngagement, recGOpe.AmjEngagement)

If currentAction = constSaisieG Or currentAction = constValider Then
    If recGOpe.AmjEngagement < paramAmjEngagementMin Then Call lstErr_AddItem(lstErr, cmdContext, "? date de l'engagement < " & dateImp(paramAmjEngagementMin))
    If recGOpe.AmjEngagement > paramAmjEngagementMax Then Call lstErr_AddItem(lstErr, cmdContext, "? date de l'engagement > " & dateImp(paramAmjEngagementMax))
End If

libEngagementCompte = ""
If Trim(txtEngagementCompte) = "" Then
    Call lstErr_AddItem(lstErr, lstErr, "? préciser la contrepartie")
    If Not blnSetfocus Then blnSetfocus = True: currentActiveControl_Name = "txtEngagementCompte"

Else
    GSub_recRacine.Numéro = CLng(num_CDec(txtEngagementCompte))
    V = srvRacineFind(GSub_recRacine)
''   MsgBox "cmdcontrol_opérationg", vbInformation, "Racine ok"
If blnJPL Then V = Null
    If Not IsNull(V) Then
        Call lstErr_AddItem(lstErr, cmdContext, "? contrepartie inconnue")
    Else
        libEngagementCompte = GSub_recRacine.Intitulé
        mCompte_Ordinaire = GSub_recRacine.Numéro & "001" & "010"
        Compte_BiaClé mCompte_Ordinaire

    End If
End If

If paramTC.NatureSens = "I" Then
    mOpération1_BackColor = crUsr.BackColor
    mOpération2_BackColor = dbUsr.BackColor
    fraOpération1.ForeColor = vbBlue
    fraOpération2.ForeColor = vbRed
Else
    mOpération1_BackColor = dbUsr.BackColor
    mOpération2_BackColor = crUsr.BackColor
    fraOpération2.ForeColor = vbBlue
    fraOpération1.ForeColor = vbRed
End If

fraOpération1.BackColor = mOpération1_BackColor
usrColor_Container fraOpération1, mOpération1_BackColor
fraOpération2.BackColor = mOpération2_BackColor
usrColor_Container fraOpération2, mOpération2_BackColor


If lstErr.ListCount = 0 And currentAction = constSaisieG Then
'    fraOpérationG.BackColor = picUsr.BackColor
'    usrColor_Container fraOpérationG, picUsr.BackColor

    currentAction = constSaisie
    fraOpérationG.Enabled = False
    fraOpération1.Enabled = True
    fraOpération2.Enabled = True
    Call cbo_Scan(paramTC.NatureDev1, cboDevise1)
    Call cbo_Scan(paramTC.NatureDev2, cboDevise2)
    GSub_CV1.DeviseIso = "": GSub_CV2.DeviseIso = ""
    recGOpe.EngagementCompte = "": recGOpe.EchéanceCompte = ""
    cmdControl_OpérationD
    blnSetfocus = True: currentActiveControl_Name = "txtMontant1"
    mGOpe = recGOpe
End If


End Sub

Public Sub cmdContext_Return()
If fraOption.Visible Then
    fraOption.Visible = False
Else
    If SSTab1.Tab = 0 Then
        'mnuList_Click
    Else
        SendKeys "{TAB}"
    End If
End If

End Sub



Public Sub cmdControl_OpérationD()
Dim I1 As Integer, I2 As Integer

Call DTPicker_Control(txtAMJValeur, recGOpe.AmjDébut)
recGOpe.AmjFin = recGOpe.AmjDébut
If recGOpe.AmjEngagement > recGOpe.AmjDébut Then Call lstErr_AddItem(lstErr, cmdContext, "? date d'engagement < date de valeur ")

Call DTPicker_Control(txtAMJEchéance, recGOpe.AmjEchéance1)
If recGOpe.AmjEngagement > recGOpe.AmjEchéance1 Then Call lstErr_AddItem(lstErr, cmdContext, "? date d'engagement < date d'échéance ")

Call cbo_Value(recGOpe.Devise1, cboDevise1)
If recGOpe.Devise1 <> xGOpe.Devise1 Then
    GSub_CV1.DeviseIso = recGOpe.Devise1
    V = CV_Attribut(GSub_CV1): If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, "? erreur devise : " & GSub_CV1.DeviseIso)
    Call srvGSub.Correspondant_cbo(cboCorrespondantN1, paramTC, GSub_CV1.DeviseN, GSub_CV1.DeviseIso, mCompte_Ordinaire)
    recGOpe.EngagementCompte = ""
End If

Call cbo_Value(recGOpe.Devise2, cboDevise2)
If recGOpe.Devise2 <> xGOpe.Devise2 Then
    GSub_CV2.DeviseIso = recGOpe.Devise2
    V = CV_Attribut(GSub_CV2): If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, "? erreur devise : " & GSub_CV2.DeviseIso)
    Call srvGSub.Correspondant_cbo(cboCorrespondantN2, paramTC, GSub_CV2.DeviseN, GSub_CV2.DeviseIso, mCompte_Ordinaire)
    recGOpe.EchéanceCompte = ""
End If

If recGOpe.Devise1 = recGOpe.Devise2 Then Call lstErr_AddItem(lstErr, lstErr, " ? Devise 1 = devise 2")


If recGOpe.EngagementCompte <> xGOpe.EngagementCompte Or Trim(recGOpe.EngagementCompte) = "" Then
    Call srvGSub_TC.paramTC_Nature_TypeDeCompte(paramTC, recGOpe, "CptàR")
    recGOpe.EngagementCompte = GSub_recRacine.Numéro & paramTC.BiatypEngagement & "010"
    Compte_BiaClé recGOpe.EngagementCompte
    libCompte1 = Compte_Imp(recGOpe.EngagementCompte)
    V = Compte_Load(GSub_CV1.DeviseN, recGOpe.EngagementCompte)
    If Not IsNull(V) Then
        recGOpe.EngagementCompte = ""
        libCompte1.ForeColor = errUsr.ForeColor
    Else
        libCompte1.ForeColor = libUsr.ForeColor
    End If
    
End If


If recGOpe.EchéanceCompte <> xGOpe.EchéanceCompte Or Trim(recGOpe.EchéanceCompte) = "" Then
    Call srvGSub_TC.paramTC_Nature_TypeDeCompte(paramTC, recGOpe, "CptàL")
    recGOpe.EchéanceCompte = GSub_recRacine.Numéro & paramTC.BiatypEngagement & "010"
    Compte_BiaClé recGOpe.EchéanceCompte
    libCompte2 = Compte_Imp(recGOpe.EchéanceCompte)
    V = Compte_Load(GSub_CV2.DeviseN, recGOpe.EchéanceCompte)
    If Not IsNull(V) Then
        recGOpe.EchéanceCompte = ""
        libCompte2.ForeColor = errUsr.ForeColor
    Else
        libCompte2.ForeColor = libUsr.ForeColor
    End If
End If

wCotation1_2 = " ": wCotation2_1 = " "
V = cmdControl_Cotation(recGOpe.Devise1, recGOpe.Devise2, wCotation1_2, wCotation2_1)
If IsNull(V) Then
    libTaux = "? cotation inconnue : "
    Call lstErr_AddItem(lstErr, lstErr, libTaux & recGOpe.Devise1 & "=" & recGOpe.Devise2)
Else
    recGOpe.TauxRéférence1 = wCotation1_2
    libTaux = V
End If


If Trim(txtMontant1) = "" Then
    Call lstErr_AddItem(lstErr, lstErr, " ? préciser le montant1")
    If Not blnSetfocus Then blnSetfocus = True: currentActiveControl_Name = "txtMontant1"
Else
    
   recGOpe.Montant1 = CCur(num_CDec(txtMontant1))
End If

If Trim(txtMontant2) = "" Then
    Call lstErr_AddItem(lstErr, lstErr, " ? préciser le montant2")
    If Not blnSetfocus Then blnSetfocus = True: currentActiveControl_Name = "txtMontant1"
Else
    
   recGOpe.Montant2 = CCur(num_CDec(txtMontant2))

End If

If Trim(txtTaux) = "" Then
    recGOpe.TauxMarge1 = 1
    Call lstErr_AddItem(lstErr, lstErr, " ? préciser le cours")
    If Not blnSetfocus Then blnSetfocus = True: currentActiveControl_Name = "txtTaux"
Else
    recGOpe.TauxMarge1 = Round(CDbl(num_CDec(txtTaux)), 5)
    txtTaux = Trim(Format$(recGOpe.TauxMarge1, "### ##0.00000"))
End If

recGOpe.Flag1 = chkMontant2

If chkMontant2 <> "1" Then
    txtMontant1.Enabled = True: txtMontant2.Enabled = False
    GSub_CV1.Montant = recGOpe.Montant1
    GSub_CV1.Cours = 1
    GSub_CV2.Cours = recGOpe.TauxMarge1
    V = CV_Manuel(GSub_CV1, GSub_CV2, GSub_CV3, X, wCotation1_2)
    recGOpe.Montant2 = GSub_CV2.Montant
Else
    txtMontant2.Enabled = True: txtMontant1.Enabled = False
    GSub_CV2.Montant = recGOpe.Montant2
    GSub_CV2.Cours = 1
    GSub_CV1.Cours = recGOpe.TauxMarge1
    V = CV_Manuel(GSub_CV2, GSub_CV1, GSub_CV3, X, wCotation2_1)
    recGOpe.Montant1 = GSub_CV1.Montant
End If

    
If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, V)

If recGOpe.Montant1 = 0 Then
    txtMontant1 = ""
Else
    txtMontant1 = Format$(recGOpe.Montant1, "### ### ### ###.00")
End If

If recGOpe.Montant2 = 0 Then
    txtMontant2 = ""
Else
    txtMontant2 = Format$(recGOpe.Montant2, "### ### ### ###.00")
End If

If Trim(recGOpe.Nature) = "CT" Then cmdControl_CT

X = cboCorrespondantN1.Text
I1 = InStr(1, X, "_")
I2 = InStr(1, X, "[")
recGOpe.EngagementCorrSwiftN = mId$(X, I1 + 1, I2 - I1)
recGOpe.EngagementCorrCompte = mId$(X, I2 + 1, 11)
If recGOpe.EngagementCorrCompte <> xGOpe.EngagementCorrCompte Then

    If cboCorrespondantN1.ListIndex <> 0 Then
        txtCorrespondantL1.Enabled = True
        If txtCorrespondantL1 = SocBicId Then txtCorrespondantL1 = ""
    Else
        txtCorrespondantL1.Enabled = False
        txtCorrespondantL1 = SocBicId
        V = Compte_Load(GSub_CV1.DeviseN, recGOpe.EngagementCorrCompte)
        If Not IsNull(V) Then
            recGOpe.EngagementCorrCompte = ""
            txtCorrespondantL1.ForeColor = errUsr.ForeColor
        Else
            txtCorrespondantL1.ForeColor = libUsr.ForeColor
        End If
    End If
End If

recGOpe.EngagementCorrSwiftL = Trim(txtCorrespondantL1)


X = cboCorrespondantN2.Text
I1 = InStr(1, X, "_")
I2 = InStr(1, X, "[")
recGOpe.EchéanceCorrSwiftN = mId$(X, I1 + 1, I2 - I1)
recGOpe.EchéanceCorrCompte = mId$(X, I2 + 1, 11)
If recGOpe.EchéanceCorrCompte <> xGOpe.EchéanceCorrCompte Then

    If cboCorrespondantN2.ListIndex <> 0 Then
        txtCorrespondantL2.Enabled = True
        If txtCorrespondantL2 = SocBicId Then txtCorrespondantL2 = ""
    Else
        txtCorrespondantL2.Enabled = False
        txtCorrespondantL2 = SocBicId
        V = Compte_Load(GSub_CV2.DeviseN, recGOpe.EchéanceCorrCompte)
        If Not IsNull(V) Then
            recGOpe.EchéanceCorrCompte = ""
            txtCorrespondantL2.ForeColor = errUsr.ForeColor
        Else
            txtCorrespondantL2.ForeColor = libUsr.ForeColor
        End If
    End If
End If

recGOpe.EchéanceCorrSwiftL = Trim(txtCorrespondantL2)

            
'If lstErr.ListCount <> 0 Then Exit Sub
End Sub

Public Sub cmdReset_Opération(lNature As String)
Select Case Trim(lNature)
    Case "CC":
        fraOpération1.Caption = "Devise achetée "
        lblTaux.Caption = "Cours"
        fraOpération2.Caption = "Devise vendue"
        lblIntérêts.Visible = False
        txtIntérêts.Visible = False
        txtAMJEchéance.Enabled = False
        txtMontant1.Enabled = True
        txtMontant2.Enabled = False: chkMontant2.Value = "0"
        chkCorrespondantL2.Enabled = False
        chkCorrespondantN2.Enabled = False
    Case "CT":
        fraOpération1.Caption = "Devise achetée "
        lblTaux.Caption = "Cours terme"
        fraOpération2.Caption = "Devise vendue"
        lblIntérêts.Visible = True: lblIntérêts.Caption = "Cours spot"
        txtIntérêts.Visible = True
        txtAMJEchéance.Enabled = False
        txtMontant1.Enabled = True
        txtMontant2.Enabled = False: chkMontant2.Value = "0"
        chkCorrespondantL2.Enabled = False
        chkCorrespondantN2.Enabled = False
End Select

End Sub

Public Function cmdControl_Cotation(lDevise1 As String, lDevise2 As String, lCotation1_2 As String, lCotation2_1 As String)


xElpTable.Method = "Seek="
xElpTable.Id = paramTC.TableId
xElpTable.K1 = "Cotation"
xElpTable.K2 = lDevise1 & "=*" & lDevise2
If tableElpTable_Read(xElpTable) = 0 Then
    lCotation1_2 = "*": lCotation2_1 = "/"
    cmdControl_Cotation = xElpTable.Name
Else
    xElpTable.K2 = lDevise2 & "=*" & lDevise1
    If tableElpTable_Read(xElpTable) = 0 Then
        lCotation1_2 = "/": lCotation2_1 = "*"
        cmdControl_Cotation = xElpTable.Name
    Else
        lCotation1_2 = "": lCotation2_1 = ""
        cmdControl_Cotation = Null
    End If
End If

    
End Function

Public Function cmdSave_Db()
If lstErr.ListCount = 0 Then
    blnControl = False
    V = srvGOpe_Update(recGOpe)
    xGOpe = recGOpe
    
    If IsNull(V) Then V = GEch_Save(recGOpe.IdRéférence)
    If IsNull(V) Then V = GFlux_Save(recGOpe.IdRéférence)
    If IsNull(V) Then V = GMemo_Save(recGOpe.IdRéférence)
    If IsNull(V) And blnComptaAuto Then
        meGMemo(1).Method = constCompta
        meGMemo(1).MemoSéquence = 0
        V = srvGMemo_Update(meGMemo(1))
    End If

cmdSave_Db = V

    If IsNull(V) Then
        If blnfgSelect_DisplayLine Then
            meGOpe(meGOpe_Index) = recGOpe
            If recGOpe.Method = constDelete Then
                fgSelect_Display
            Else
                fgSelect_DisplayLine
            End If
        End If
        lastActiveControl_Name = ""
        cmdOk.Visible = False
        cmdSave.Visible = False
        Call lstErr_Clear(lstErr, cmdContext, "Mise à jour effectuée - Identification : " & recGOpe.IdRéférence)
    ''    cmdContext_Quit
    Else
        Call lstErr_Clear(lstErr, cmdContext, V)
 ''''       cmdReset
    End If

End If

End Function

Public Sub GEch_GenCC()
On Error GoTo Error_Handle

'If recGOpe.Method = constAddNew Then
recGEch = meGECh(1)

    recGEch.Method = constAddNew
    recGEch.EchAMJ = recGOpe.AmjEngagement
    recGEch.EchHMS = "000000"
    recGEch.EchUsr = "Auto"
    If recGOpe.AmjEngagement = DSys And TCAut.Comptabiliser Then
        recGEch.EchUsr = ""
        recGEch.Statut = "@"
        recGEch.StatutPlus = "C"
    End If
    
    With recGEch                                   ' Compta HB  1
        .EchSéquence = recGEch.EchSéquence + 1
        .FluxSéquence = 1
        .EchFct = constComptaHB
    End With
    meGECh(recGEch.EchSéquence) = recGEch
    
    With recGEch                                   ' Compta HB  2
        .EchSéquence = recGEch.EchSéquence + 1
       .FluxSéquence = 2
       .EchFct = constComptaHB
    End With
    meGECh(recGEch.EchSéquence) = recGEch
    
    recGEch.EchUsr = "Auto"
    recGEch.Statut = ""
    recGEch.StatutPlus = ""
    
   With recGEch                                   ' Compta 1
        .EchSéquence = recGEch.EchSéquence + 1
        .FluxSéquence = 1
        .EchFct = constCompta
        .EchAMJ = recGOpe.AmjDébut
    End With
    meGECh(recGEch.EchSéquence) = recGEch
     
    With recGEch                                   ' Compta 1
        .EchSéquence = recGEch.EchSéquence + 1
        .FluxSéquence = 2
        .EchFct = constCompta
        .EchAMJ = recGOpe.AmjDébut
    End With
    meGECh(recGEch.EchSéquence) = recGEch
    
    If recGOpe.EngagementCorrSwiftL <> SocBicId Then
        With recGEch                                   ' Swift  1
            .EchSéquence = recGEch.EchSéquence + 1
            .FluxSéquence = 1
            .EchFct = constSwiftSnd
            .EchAMJ = recGOpe.AmjDébut
        End With
        meGECh(recGEch.EchSéquence) = recGEch
    End If
    If recGOpe.EchéanceCorrSwiftL <> SocBicId Then
        With recGEch                                   ' Swift 2
            .EchSéquence = recGEch.EchSéquence + 1
            .FluxSéquence = 2
            .EchFct = constSwiftSnd
            .EchAMJ = recGOpe.AmjDébut
        End With
        meGECh(recGEch.EchSéquence) = recGEch
    End If
    
    meGECh_Nb = recGEch.EchSéquence
    
'End If


Exit Sub
'---------------------------------------------------------
Error_Handle:
'---------------------------------------------------------

Call MsgBox("Erreur", vbCritical, "gech_GenCC")
End Sub

Public Sub GEch_Gen()

Select Case Trim(recGOpe.Nature)
    Case "CC": GEch_GenCC
    Case "CT": GEch_GenCT
End Select

End Sub

Public Sub mnuOpérationSaisir_Init()
cmdReset
SSTab1.Tab = 2
fgSelect.Enabled = False
fgFlux.Enabled = False
fgFlux.Clear: fgFlux.Rows = 1: fgFlux_RowDisplay = 0
fraOpérationG.Enabled = True
fraOpération1.Enabled = False
fraOpération2.Enabled = False
fraOpération.Enabled = True

blncmdOk_Visible = True: blncmdSave_Visible = True
blnAmjEchéance = False
txtRéférenceInterne.SetFocus
cmdContext.Caption = constcmdAbandonner

GEch_GenNew

blnControl = True

End Sub

Public Sub GEch_GenNew()

ReDim meGECh(10)

srvGEch.recGEch_Init recGEch
recGEch.Method = constAddNew

With recGEch                                   ' Saisie
    .IdRéférence = recGOpe.IdRéférence
    .EchSéquence = 1
    .Application = paramTC.Application
    .EchFct = constSaisie
    .EchAMJ = DSys
    .EchHMS = time_Hms
    .EchUsr = usrId
End With
meGECh(recGEch.EchSéquence) = recGEch

End Sub

Public Sub Gech_GenUpdate()

If Trim(meGECh(1).Method) = "" Then meGECh(1).Method = constUpdate

With meGECh(1)                                  ' Saisie
    .EchAMJ = DSys
    .EchHMS = time_Hms
    .EchUsr = usrId
End With

End Sub

Public Function GFlux_Scan(lFluxSéquence As Long)
GFlux_Scan = "?GFlux_Scan"
For meGFlux_Index = 1 To meGFlux_Nb
    If meGFlux(meGFlux_Index).FluxSéquence = lFluxSéquence Then
        recGFlux = meGFlux(meGFlux_Index)
        GFlux_Scan = Null
        Exit For
    End If
Next meGFlux_Index

End Function
Public Function GMemo_Scan(lEchSéquence As Long)
GMemo_Scan = "?GMemo_Scan"
wGMemo_Nb = 0
For meGMemo_Index = 1 To meGMemo_Nb
    If meGMemo(meGMemo_Index).EchSéquence = lEchSéquence Then
        If wGMemo_Nb = wGMemo_NbMax Then wGMemo_NbMax = wGMemo_NbMax + 10: ReDim Preserve wGMemo(wGMemo_NbMax)
        wGMemo_Nb = wGMemo_Nb + 1
        wGMemo(wGMemo_Nb) = meGMemo(meGMemo_Index)
    End If
Next meGMemo_Index
End Function


Public Function GEch_Scan(lFluxSéquence As Long)
GEch_Scan = "?GEch_Scan"
wGECh_Nb = 0
ReDim wGEch(meGECh_Index + 1)
For meGECh_Index = 1 To meGECh_Nb
    If meGECh(meGECh_Index).FluxSéquence = lFluxSéquence Then
        wGECh_Nb = wGECh_Nb + 1
        wGEch(wGECh_Nb) = meGECh(meGECh_Index)
    End If
Next meGECh_Index
End Function

Public Sub cmdOk_Valider()
Dim I As Integer, K As Integer

recGOpe.Statut = " "
recGOpe.StatutPlus = "  "
With meGECh(1)                                  ' Saisie
    .Method = constUpdate
    .ActionFct = constValider
    .ActionAmj = DSys
    .ActionHms = time_Hms
    .ActionUsr = usrId
    .Statut = "F"
    .StatutPlus = "in"
End With

meGMemo_Nb = 0

For I = 1 To meGECh_Nb
    If meGECh(I).Statut = "@" And meGECh(I).StatutPlus = "C " Then
        recGEch = meGECh(I)
        V = GFlux_Scan(recGEch.FluxSéquence)
        If IsNull(V) Then
            With meGECh(I)
                .ActionFct = constCompta
                .ActionAmj = DSys
                .ActionHms = time_Hms
                .ActionUsr = usrId
            End With
            Call srvGSub_TC.GMemo_Gen(paramTC, recGOpe, recGFlux, recGEch, wGMemo_Nb, wGMemo())
            For K = 1 To wGMemo_Nb
                If meGMemo_Nb = meGMemo_NbMax Then meGMemo_NbMax = meGMemo_NbMax + 10: ReDim Preserve meGMemo(meGMemo_NbMax)

                meGMemo_Nb = meGMemo_Nb + 1
                meGMemo(meGMemo_Nb) = wGMemo(K)
                meGMemo(meGMemo_Nb).Method = constAddNew
                meGMemo(meGMemo_Nb).MemoSéquence = meGMemo_Nb
                meGMemo(meGMemo_Nb).MemoLien1 = paramTC.ComptaLot
                meGMemo(meGMemo_Nb).Statut = "@"
                meGMemo(meGMemo_Nb).StatutPlus = "C "
            Next K
        End If
    End If
Next I

If meGMemo_Nb > 0 Then blnComptaAuto = TCAut.Comptabiliser

End Sub


Public Sub GEch_GenCT()
GEch_GenCC
End Sub

Public Sub cmdControl_CT()
Dim X As String
If Trim(txtIntérêts) = "" Then
    recGOpe.TauxMarge2 = 1
    Call lstErr_AddItem(lstErr, lstErr, " ? préciser le cours terme")
    If Not blnSetfocus Then blnSetfocus = True: currentActiveControl_Name = "txtTaux"
Else
    recGOpe.TauxMarge2 = Round(CDbl(num_CDec(txtIntérêts)), 5)
    txtIntérêts = Trim(Format$(recGOpe.TauxMarge2, "### ##0.00000"))
End If

GSub_CV1.Montant = recGOpe.Montant1
GSub_CV1.Cours = 1
GSub_CV2.Cours = recGOpe.TauxMarge2
V = CV_Manuel(GSub_CV1, GSub_CV2, GSub_CV3, X, wCotation1_2)
recGOpe.Mensualité = GSub_CV2.Montant

If recGOpe.TauxMarge1 > recGOpe.TauxMarge2 Then
        X = "Report       : "
Else
        X = "Déport       : "
End If

libInfo = recGOpe.Devise2 & " Terme  : " & Format$(recGOpe.Montant2, "### ### ### ###.00") & Chr$(13) _
        & recGOpe.Devise2 & " Spot     : " & Format$(recGOpe.Mensualité, "### ### ### ###.00") & Chr$(13) & Chr$(13) _
        & X & Format$(recGOpe.TauxMarge1 - recGOpe.TauxMarge2, "###.00000")

End Sub
