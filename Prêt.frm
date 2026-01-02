VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrêt 
   AutoRedraw      =   -1  'True
   Caption         =   "Prêts"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6765
   ScaleWidth      =   9420
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
      Height          =   500
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   0
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
      Height          =   500
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   6120
      TabIndex        =   17
      Top             =   0
      Width           =   2745
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8880
      Picture         =   "Prêt.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   0
      Width           =   500
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
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   0
      TabIndex        =   18
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Sélection"
      TabPicture(0)   =   "Prêt.frx":0102
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Liste des prêts"
      TabPicture(1)   =   "Prêt.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgSelect"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Prêt "
      TabPicture(2)   =   "Prêt.frx":013A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraPrêt"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tableau d'amortissement"
      TabPicture(3)   =   "Prêt.frx":0156
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fgEchéancier"
      Tab(3).ControlCount=   1
      Begin VB.Frame fraSelect 
         Caption         =   "Critères de sélection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   -74760
         TabIndex        =   29
         Top             =   840
         Width           =   8775
         Begin VB.TextBox txtSelect 
            Height          =   285
            Left            =   3000
            TabIndex        =   34
            Top             =   960
            Width           =   3615
         End
         Begin VB.OptionButton optSelectRéférenceExterne 
            Caption         =   "Référence externe"
            Height          =   375
            Left            =   360
            TabIndex        =   33
            Top             =   1920
            Width           =   1935
         End
         Begin VB.OptionButton optSelectRéférenceInterne 
            Caption         =   "Régérence interne"
            Height          =   255
            Left            =   360
            TabIndex        =   32
            Top             =   1440
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton optSelectCompte 
            Caption         =   "Compte"
            Height          =   255
            Left            =   360
            TabIndex        =   31
            Top             =   2520
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker txtAmjMax 
            Height          =   300
            Left            =   3000
            TabIndex        =   36
            Top             =   3120
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
            Format          =   24576003
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin MSComCtl2.DTPicker txtAmjMin 
            Height          =   300
            Left            =   5520
            TabIndex        =   37
            Top             =   3120
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
            Format          =   24576003
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblSelect2 
            Caption         =   "Echéancier jusqu'au :"
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
            Left            =   360
            TabIndex        =   35
            Top             =   3120
            Width           =   2295
         End
         Begin VB.Label lblSelect1 
            Caption         =   "Afficher la liste des prêts"
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
            Left            =   360
            TabIndex        =   30
            Top             =   960
            Width           =   1935
         End
      End
      Begin VB.Frame fraPrêt 
         Caption         =   "Caractéristiques du prêt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   0
         TabIndex        =   20
         Top             =   360
         Width           =   9135
         Begin VB.Frame fraRésultat 
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
            Height          =   2055
            Left            =   5520
            TabIndex        =   53
            Top             =   240
            Width           =   3375
            Begin VB.TextBox txtMensualité 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               TabIndex        =   57
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox txtTEG 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               TabIndex        =   56
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox txtTauxActuariel 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               TabIndex        =   55
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox txtIdRéférence 
               Height          =   285
               Left            =   1680
               TabIndex        =   54
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label lblIdRéférence 
               Caption         =   "Identification"
               Height          =   255
               Left            =   240
               TabIndex        =   61
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label lblTauxActuariel 
               Caption         =   "Taux Actuariel"
               Height          =   255
               Left            =   240
               TabIndex        =   60
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label lblTEG 
               Caption         =   "T E G"
               Height          =   255
               Left            =   240
               TabIndex        =   59
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label lblMensualité 
               Caption         =   "Mensualité"
               Height          =   255
               Left            =   240
               TabIndex        =   58
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame fraCompta 
            Caption         =   "Comptabilisation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3255
            Left            =   120
            TabIndex        =   38
            Top             =   2280
            Width           =   8895
            Begin VB.TextBox txtCorrCompte 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1560
               TabIndex        =   11
               Top             =   2880
               Width           =   1575
            End
            Begin VB.TextBox txtEngCompte 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1560
               TabIndex        =   10
               Top             =   2400
               Width           =   1575
            End
            Begin VB.TextBox txtEchCompte 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1560
               TabIndex        =   9
               Top             =   1920
               Width           =   1575
            End
            Begin VB.CheckBox chkComptaReprise 
               Caption         =   "Reprise "
               Height          =   255
               Left            =   7800
               TabIndex        =   41
               Top             =   2880
               Width           =   975
            End
            Begin VB.ComboBox cboNature 
               Height          =   315
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   480
               Width           =   4815
            End
            Begin VB.OptionButton optEchéanceFinDeMois 
               Caption         =   "Fin de mois"
               Height          =   195
               Left            =   6000
               TabIndex        =   40
               Top             =   1080
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton optEchéanceAnniversaire 
               Caption         =   "Anniversaire"
               Height          =   195
               Left            =   7200
               TabIndex        =   39
               Top             =   1080
               Width           =   1215
            End
            Begin VB.TextBox txtRéférenceInterne 
               Height          =   285
               Left            =   1560
               TabIndex        =   7
               Top             =   1440
               Width           =   2655
            End
            Begin VB.TextBox txtRéférenceExterne 
               Height          =   285
               Left            =   5880
               TabIndex        =   8
               Top             =   1440
               Width           =   2655
            End
            Begin MSComCtl2.DTPicker txtAmjEngagement 
               Height          =   300
               Left            =   1560
               TabIndex        =   5
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
               Format          =   24576003
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtAmjEchéance 
               Height          =   300
               Left            =   4560
               TabIndex        =   6
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
               Format          =   24576003
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label libCorrCompte 
               Caption         =   "-"
               Height          =   255
               Left            =   3480
               TabIndex        =   52
               Top             =   2880
               Width           =   4095
            End
            Begin VB.Label lblCorrCompte 
               Caption         =   "Déboclage des fonds"
               Height          =   375
               Left            =   120
               TabIndex        =   51
               Top             =   2760
               Width           =   1095
            End
            Begin VB.Label libEngCompte 
               Caption         =   "-"
               Height          =   255
               Left            =   3480
               TabIndex        =   50
               Top             =   2400
               Width           =   5055
            End
            Begin VB.Label lblEngCompte 
               Caption         =   "Compte de prêt"
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   2400
               Width           =   1335
            End
            Begin VB.Label libEchCompte 
               Caption         =   "-"
               Height          =   255
               Left            =   3480
               TabIndex        =   48
               Top             =   1920
               Width           =   5175
            End
            Begin VB.Label lblEchCompte 
               Caption         =   "Racine /Compte  à prélever"
               Height          =   495
               Left            =   120
               TabIndex        =   47
               Top             =   1800
               Width           =   1335
            End
            Begin VB.Label lblNature 
               Caption         =   "Nature"
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label lblAmjDébut 
               Caption         =   "Date du prêt."
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lblAmjEchéance 
               Caption         =   "Première échéance"
               Height          =   255
               Left            =   2880
               TabIndex        =   44
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label lblRéférenceInterne 
               Caption         =   "Référence interne"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   1440
               Width           =   1575
            End
            Begin VB.Label lblRéférenceExterne 
               Caption         =   "Référence externe"
               Height          =   255
               Left            =   4440
               TabIndex        =   42
               Top             =   1440
               Width           =   1575
            End
         End
         Begin VB.Frame fraNature 
            Caption         =   "(Remboursement constant)"
            Height          =   2055
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   5295
            Begin VB.OptionButton optAnnuel 
               Caption         =   "An"
               Height          =   255
               Left            =   4440
               TabIndex        =   64
               Top             =   1560
               Width           =   735
            End
            Begin VB.OptionButton optSemestriel 
               Caption         =   "Sem"
               Height          =   255
               Left            =   3600
               TabIndex        =   63
               Top             =   1560
               Width           =   735
            End
            Begin VB.OptionButton optTrimestriel 
               Caption         =   "Trim"
               Height          =   255
               Left            =   2880
               TabIndex        =   62
               Top             =   1560
               Width           =   735
            End
            Begin VB.TextBox txtDevise 
               Height          =   285
               Left            =   1320
               TabIndex        =   12
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox txtCapital 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1320
               TabIndex        =   0
               Top             =   840
               Width           =   1575
            End
            Begin VB.TextBox txtTaux 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1320
               TabIndex        =   2
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox txtPériodeNb 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1320
               TabIndex        =   3
               Top             =   1560
               Width           =   735
            End
            Begin VB.OptionButton optMensuel 
               Caption         =   "Mois"
               Height          =   255
               Left            =   2160
               TabIndex        =   22
               Top             =   1560
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.TextBox txtFrais 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3600
               TabIndex        =   1
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label lblDevise 
               Caption         =   "Devise"
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   360
               Width           =   615
            End
            Begin VB.Label lblCapital 
               Caption         =   "Capital"
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   840
               Width           =   615
            End
            Begin VB.Label lblTaux 
               Caption         =   "Taux annuel"
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label lblPériodeNB 
               Caption         =   "Nb périodes"
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label lblFrais 
               Caption         =   "Frais"
               Height          =   255
               Left            =   3000
               TabIndex        =   23
               Top             =   960
               Width           =   615
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgEchéancier 
         Height          =   5250
         Left            =   -74880
         TabIndex        =   19
         Top             =   600
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   9260
         _Version        =   393216
         Rows            =   1
         Cols            =   12
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
         FormatString    =   $"Prêt.frx":0172
      End
      Begin MSFlexGridLib.MSFlexGrid fgSelect 
         Height          =   5250
         Left            =   -74880
         TabIndex        =   28
         Top             =   600
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   9260
         _Version        =   393216
         Rows            =   1
         Cols            =   12
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
         FormatString    =   $"Prêt.frx":02B8
      End
   End
   Begin MSComCtl2.DTPicker txtAMJ 
      Height          =   300
      Left            =   3840
      TabIndex        =   65
      Top             =   120
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
      Format          =   24444931
      CurrentDate     =   36299
      MaxDate         =   401768
      MinDate         =   -328351
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuPrêtSaisir 
         Caption         =   "Saisir un prêt"
      End
      Begin VB.Menu mnuContextX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListàValider 
         Caption         =   "Liste des prêts à valider"
      End
      Begin VB.Menu mnuListPrêts 
         Caption         =   "Liste des prêts"
      End
      Begin VB.Menu mnuListEchéancier 
         Caption         =   "Echéancier"
      End
      Begin VB.Menu mnuContextX3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComptaEchéancier 
         Caption         =   "Compta : Echéances à comptabiliser = Jour"
      End
      Begin VB.Menu mnuComptaEchéancier_Plus 
         Caption         =   "Compta : Echéances à comptabiliser > Jour"
      End
      Begin VB.Menu mnuComptaLotsàValider 
         Caption         =   "Compta : Lots à valider"
      End
      Begin VB.Menu mnuLotComptabilisé_Annuler 
         Caption         =   "Compta : annuler un lot comptabilisé"
      End
      Begin VB.Menu mnuContextX2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuPrêt 
      Caption         =   "mnuPrêt"
      Visible         =   0   'False
      Begin VB.Menu mnuContextX4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrêtDisplay 
         Caption         =   "Afficher ce prêt"
      End
      Begin VB.Menu mnuPrêtModifier 
         Caption         =   "Modifier ce prêt"
      End
      Begin VB.Menu mnuPrêtValider 
         Caption         =   "Valider/ Invalider ce prêt"
      End
      Begin VB.Menu mnuPrêtRemboursementAnticipé 
         Caption         =   "Remboursement anticipé"
      End
      Begin VB.Menu mnuPrêtAnnuler 
         Caption         =   "Annuler ce prêt"
      End
      Begin VB.Menu mnuPrêtModifierEchCompte 
         Caption         =   "Changer le compte à prélever"
      End
      Begin VB.Menu mnuPrêtBasculeEuro 
         Caption         =   "Bascule Euro"
      End
      Begin VB.Menu mnuContextX5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrêtEffacer 
         Caption         =   "Effacer ce prêt"
      End
      Begin VB.Menu mnuPrêtPrint 
         Caption         =   "Imprimer ce prêt"
      End
      Begin VB.Menu mnuPrêtPrintList 
         Caption         =   "Imprimer la liste sélectionnée"
      End
   End
   Begin VB.Menu mnucmdPrint 
      Caption         =   "mnucmdPrint"
      Visible         =   0   'False
      Begin VB.Menu mnucmdPrintTableauAmortissement 
         Caption         =   "Imprimer le tableau d'amortissement"
      End
      Begin VB.Menu mnucmdPrintPrêt 
         Caption         =   "Imprimer le prêt sélectionné"
      End
      Begin VB.Menu mnucmdPrintList_TOpe 
         Caption         =   "Imprimer la liste des prêts sélectionnés"
      End
      Begin VB.Menu mnucmdPrintList_TFlux 
         Caption         =   "Imprimer l'échéancier"
      End
   End
   Begin VB.Menu mnuLot 
      Caption         =   "mnuLot"
      Visible         =   0   'False
      Begin VB.Menu mnuLotàComptaValidation 
         Caption         =   "Lot : comptabilisation définitive"
      End
      Begin VB.Menu mnuLotàComptaAnnulation 
         Caption         =   "Lot : annuler la demande de comptabilisation"
      End
      Begin VB.Menu mnuLotàComptaPrint 
         Caption         =   "Lot : imprimer la demande de comptabilisation"
      End
   End
End
Attribute VB_Name = "frmPrêt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim PrêtsAut As typeAuthorization

Dim recTable As typeElpTable
Dim wAmjEngagement As String, wAmjEchéance As String, blnAmjEchéance As Boolean
Dim paramAmjEngagementMin As String, paramAmjEngagementMax As String
Dim paramAmjEchéanceMin As String, paramAmjEchéanceMax As String
''Dim fgSelect_FormatString As String, fgSelect_K As Integer
''Dim fgEchéancier_FormatString As String, fgEchéancier_K As Integer

Dim CV1 As typeCV

Dim recTope As typeTOpe, xTOpe As typeTOpe, mTope As typeTOpe
Dim arrTFlux() As typeTFlux, recTFlux As typeTFlux, mTflux As typeTFlux
Dim arrTFlux_Nb As Integer, arrTFlux_Index As Integer

Dim totalCapital As Currency, totalIntérêts As Currency
Dim recCompte As typeCompte
Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnfgSelect_DisplayLine As Boolean, blnfgEchéance_DisplayLine As Boolean

Dim fctPériodicité As String
Dim mAmjMin As String, mAmjMax As String

Dim mEChéanceCompte As String, mEngagementCompte As String, mEngagementCorrCompte As String
Dim wAmjEchéanceTrt As String * 8, wAmjRemboursementAnticipé As String * 8, minAmjRemboursementAnticipé As String * 8
Dim mNature As String

Dim blnComptaAuto As Boolean

Dim fgEchéancier_FormatString As String, fgEchéancier_K As Integer
Dim fgEchéancier_RowDisplay As Integer, fgEchéancier_RowClick As Integer
Dim fgEchéancier_ColorClick As Long, fgEchéancier_ColorDisplay As Long
Dim fgEchéancier_Sort1 As Integer, fgEchéancier_Sort2 As Integer

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer



Public Sub fgSelect_Load()
Dim X As String, mMethod As String

recTOpe_Init xTOpe

Select Case currentAction
    Case "mnuListàValider"
            xTOpe.Method = "SnapLS"
            xTOpe.Application = paramTFlux_Service
            xTOpe.IdRéférence = "0000000000"
            xTOpe.Statut = "à"
            
            arrTOpe(0) = xTOpe
            arrTOpe(0).IdRéférence = "999999999"

    Case "mnuListPrêts"
            xTOpe.Method = "SnapLRI"
            X = Trim(txtSelect)
            xTOpe.RéférenceInterne = X
            If X = "" Then xTOpe.Method = "SnapLS"
            xTOpe.Application = paramTFlux_Service
            xTOpe.IdRéférence = 0
            xTOpe.Statut = " "
            
            arrTOpe(0) = xTOpe
            arrTOpe(0).IdRéférence = 999999999
            arrTOpe(0).RéférenceInterne = X & "99999999999"

End Select

mMethod = Trim(xTOpe.Method)
arrTOpe_NBMax = 0
arrTOpe_Suite = True: arrTOpe_NB = 0
Do Until Not arrTOpe_Suite
    srvTOpe_Monitor xTOpe
    xTOpe = arrTOpe(arrTOpe_NB)
    xTOpe.Method = mMethod & "+"
Loop
fgSelect_Display
End Sub
Private Sub fgSelect_Display()
Dim K2 As Integer, I As Integer
Dim curDB As Currency, curCR As Currency, curX As Currency

SSTab1.Tab = 1

fgSelect.Visible = True
fgSelect.Clear: fgSelect.Rows = 1

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Enabled = True
For arrTOpe_Index = 1 To arrTOpe_NB
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine
    
Next arrTOpe_Index
If fgSelect.Rows = 1 Then Exit Sub
'fgSelect_Sort

End Sub
Public Sub fgSelect_Sort()
If fgSelect.Rows > 1 Then
    fgSelect.Row = 1
    fgSelect.RowSel = fgSelect.Rows - 1
    
    fgSelect.Col = fgSelect_Sort1
    fgSelect.ColSel = fgSelect_Sort2
    fgSelect.Sort = 1
End If

End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
blnControl = False
If currentAction <> "" Then
    currentAction = ""
    cmdContext.Caption = constcmdRechercher
    fgSelect.Enabled = True
    fgEchéancier.Enabled = True
    fraSelect.Enabled = True
    fraPrêt.Enabled = False
    cmdReset
    SSTab1.Tab = 0
Else
    If blnMsgBox_Quit Then
    X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
    Else
       X = vbYes
    End If
    If X = vbYes Then Unload Me
End If

End Sub
Public Sub cmdContext_Return()
If SSTab1.Tab = 0 And Trim(txtSelect) <> "" Then
    mnuListPrêts_Click
Else
    SendKeys "{TAB}"
End If
End Sub


'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me

End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
usrColor_Set
cmdOk.Caption = constàValider: cmdOk.Visible = False
cmdSave.Caption = constEnAttente: cmdSave.Visible = False
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
'blnControl = True
blncmdOk_Visible = False: blncmdSave_Visible = False
blnfgSelect_DisplayLine = False: blnfgEchéance_DisplayLine = False

fraPrêt.Enabled = False
fgEchéancier.Clear: fgEchéancier.Rows = 1
txtIdRéférence = ""
cboNature.ListIndex = 0
CV1 = CV_Euro
CV1.DeviseIso = "EUR"
CV_Attribut CV1
txtDevise = CV1.DeviseIso
txtCapital = ""
txtTaux = ""
txtPériodeNb = ""
txtFrais = ""
txtMensualité = ""
txtTEG = ""
txtTauxActuariel = ""
optMensuel = True

wAmjEngagement = DSys
Call DTPicker_Set(txtAmjEngagement, wAmjEngagement)
wAmjEchéance = dateFinDeMois(dateElp("MoisAdd", 1, DSys))
Call DTPicker_Set(txtAMJEchéance, wAmjEchéance)

txtRéférenceInterne = ""
txtRéférenceExterne = ""
txtEngCompte = "": libEngCompte = ""
txtEchCompte = "": libEchCompte = ""
txtCorrCompte = "": libCorrCompte = ""
optEchéanceFinDeMois = True
recTOpe_Init mTope
mTope.Statut = "à"
mTope.StatutPlus = "?"
mTope.Method = constAddNew
mEChéanceCompte = Space$(11): mEngagementCompte = Space$(11): mEngagementCorrCompte = Space$(11)
txtAMJ.Visible = False
blnComptaAuto = False
blnControl = True
End Sub



Public Sub Form_Init(Msg As String)
SSTab1.Tab = 0
tableElpTable_Open
fraRésultat.Enabled = False
Call DTPicker_Set(txtAmjMin, DSys)
If mNature = "PP_" Then
    wAmjEchéanceTrt = dateElp("Jour", 2, DSys)
Else
    wAmjEchéanceTrt = dateElp("Jour", 8, DSys)
End If
Call DTPicker_Set(txtAmjMax, wAmjEchéanceTrt)
paramAmjEngagementMin = mId$(DSys, 1, 6) & "01"
paramAmjEngagementMax = dateElp("Ouvré", 7, DSys)
ReDim arrTOpe(1)
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_FormatString = fgSelect.FormatString
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0

fgEchéancier_Sort1 = 0: fgEchéancier_Sort2 = 0
fgEchéancier_FormatString = fgEchéancier.FormatString
fgEchéancier_RowDisplay = 0: fgEchéancier_RowClick = 0
cmdReset
mnuPrêtSaisir.Enabled = PrêtsAut.Saisir
mnuListàValider.Enabled = PrêtsAut.Consulter
mnuListPrêts.Enabled = PrêtsAut.Consulter
mnuComptaEchéancier.Enabled = PrêtsAut.Valider
mnuComptaEchéancier_Plus.Enabled = PrêtsAut.Valider
mnuComptaLotsàValider.Enabled = PrêtsAut.Comptabiliser
mnuLotComptabilisé_Annuler.Enabled = PrêtsAut.Xspécial
blnControl = False
txtSelect.SetFocus
End Sub

Private Sub fgEchéancier_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y <= fgEchéancier.RowHeightMin Then
    Select Case fgEchéancier.Col
        Case 0: fgEchéancier_Sort1 = 0: fgEchéancier_Sort2 = 0: fgEchéancier_Sort
        Case 1: fgEchéancier_Sort1 = 1: fgEchéancier_Sort2 = 1: fgEchéancier_Sort
        Case 2: fgEchéancier_Sort1 = 2: fgEchéancier_Sort2 = 2: fgEchéancier_Sort
        Case 3: fgEchéancier_Sort1 = 3: fgEchéancier_Sort2 = 3: fgEchéancier_Sort
        Case 6: fgEchéancier_Sort1 = 6: fgEchéancier_Sort2 = 6: fgEchéancier_Sort
        Case 7: fgEchéancier_Sort1 = 7: fgEchéancier_Sort2 = 7: fgEchéancier_Sort
        Case 8: fgEchéancier_Sort1 = 8: fgEchéancier_Sort2 = 8: fgEchéancier_Sort
        Case 11: fgEchéancier_Sort1 = 11: fgEchéancier_Sort2 = 11: fgEchéancier_Sort
    End Select
Else
    fgEchéancier_K = fgEchéancier.Row * fgEchéancier.Cols
    If fgEchéancier.Rows > 1 Then
            Call fgEchéancier_Color(fgEchéancier_RowClick, MouseMoveUsr.BackColor, fgEchéancier_ColorClick)
        arrTFlux_Index = Val(fgEchéancier.TextArray(11 + fgEchéancier_K))
        recTFlux.CptMvtLot = arrTFlux(arrTFlux_Index).CptMvtLot
        
        If recTFlux.CptMvtLot > 0 Then
            mnuLotàComptaValidation = False
            mnuLotàComptaAnnulation = False
            mnuLotàComptaAnnulation = False
          
            If recTFlux.Statut = "à" And recTFlux.StatutPlus = "C " Then
                mnuLotàComptaValidation = PrêtsAut.Comptabiliser
                mnuLotàComptaAnnulation = PrêtsAut.Comptabiliser
                mnuLotàComptaPrint = PrêtsAut.Comptabiliser
            End If
    
            Me.PopupMenu mnuLot, vbPopupMenuLeftButton
        End If
    End If
End If

End Sub
Public Sub fgEchéancier_Sort()
If fgEchéancier.Rows > 1 Then
    fgEchéancier.Row = 1
    fgEchéancier.RowSel = fgEchéancier.Rows - 1
    
    fgEchéancier.Col = fgEchéancier_Sort1
    fgEchéancier.ColSel = fgEchéancier_Sort2
    fgEchéancier.Sort = 1
End If

End Sub



Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xStatut As String
If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1: fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 5: fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_Sort
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
        Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
        Case 9: fgSelect_Sort1 = 9: fgSelect_Sort2 = 9: fgSelect_Sort
        Case 10: fgSelect_Sort1 = 10: fgSelect_Sort2 = 10: fgSelect_Sort
        Case 11: fgSelect_Sort1 = 11: fgSelect_Sort2 = 11: fgSelect_Sort
    End Select
Else

    fgSelect_K = fgSelect.Row * fgSelect.Cols
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        arrTOpe_Index = Val(fgSelect.TextArray(11 + fgSelect_K))
        xTOpe = arrTOpe(arrTOpe_Index)
    
        If xTOpe.IdRéférence > 0 Then
            mnuPrêtDisplay = PrêtsAut.Consulter
            mnuPrêtModifier = False
            mnuPrêtAnnuler = False
            mnuPrêtEffacer = False
            mnuPrêtValider = False
            mnuPrêtRemboursementAnticipé = False
            mnuPrêtModifierEchCompte = False
            mnuPrêtBasculeEuro.Enabled = False
            xStatut = xTOpe.Statut & xTOpe.StatutPlus
            If xStatut = "à? " Then
                mnuPrêtModifier = PrêtsAut.Saisir
                mnuPrêtEffacer = PrêtsAut.Saisir
            End If
            If xStatut = "àV " Then
     '           If xTOpe.MajUsr = usrId Then
      '              Call lstErr_Clear(lstErr, cmdContext, "! Vous ne pouvez pas valider vos opérations")
      '          Else
                    mnuPrêtValider = PrêtsAut.Valider
       '         End If
            End If
            If xStatut = "   " Then
                mnuPrêtRemboursementAnticipé = PrêtsAut.Valider
                mnuPrêtModifierEchCompte = PrêtsAut.Saisir
                If xTOpe.Devise = "FRF" Then mnuPrêtBasculeEuro.Enabled = PrêtsAut.Xspécial
           End If
    
            Me.PopupMenu mnuPrêt, vbPopupMenuLeftButton
        End If
    End If
End If

End Sub

Private Sub mnuPrêtBasculeEuro_Click()
Dim X As String
Dim CV1 As typeCV, CV2 As typeCV, CV3 As typeCV
Dim blnAmjEchéance1 As Boolean

'currentAction = constValider
mnuPrêtSaisir_Click
fraPrêt.Enabled = False
fraPrêt_Load "Update"
blnAmjEchéance1 = False
 blncmdSave_Visible = False
For arrTFlux_Index = 1 To arrTFlux_Nb
    recTFlux = arrTFlux(arrTFlux_Index)
    If recTFlux.CodeOpération = "PR02" Then
        If recTFlux.Statut <> " " Then
            mTope.Capital = mTope.Capital - recTFlux.Capital
            mTope.PériodeNb = mTope.PériodeNb - 1
            mTope.AmjDébut = recTFlux.AmjFin
        Else
            If Not blnAmjEchéance1 Then blnAmjEchéance1 = True: mTope.AmjEchéance1 = recTFlux.AmjFin
       End If
        
    End If
Next arrTFlux_Index

If mTope.Capital <= 0 Or mTope.PériodeNb <= 0 Then
    Call MsgBox("Capital = 0", vbInformation, "Bascule Euro : PRETS")
    Exit Sub
End If

CV1 = CV_Euro: CV2 = CV_Euro: CV3 = CV_Euro

CV1.DeviseIso = "FRF": CV1.Montant = mTope.Capital
CV2.DeviseIso = "EUR"
Call CV_Transitoire(CV1, CV2, CV3, X)
mTope.Capital = CV2.Montant

mTope.Devise = "EUR"
mTope.optReprise = "R"
mTope.RéférenceInterne = Trim(mTope.RéférenceInterne) & "_EUR"
mTope.IdRéférenceLiée = mTope.IdRéférence
mTope.IdRéférence = 0
mTope.MajUsr = "EURO"
mTope.MajAMJ = DSys
mTope.Statut = "à"
mTope.StatutPlus = "V "
mTope.MajHMS = time_Hms

txtIdRéférence = ""
txtDevise = mTope.Devise
txtCapital = Format$(mTope.Capital, "### ### ### ##0.00")
txtPériodeNb = Format$(mTope.PériodeNb, "###0")
txtMensualité = Format$(mTope.Mensualité, "### ### ### ##0.00")
txtRéférenceInterne = Trim(mTope.RéférenceInterne)
Call DTPicker_Set(txtAmjEngagement, mTope.AmjDébut): wAmjEngagement = mTope.AmjDébut
Call DTPicker_Set(txtAMJEchéance, mTope.AmjEchéance1): wAmjEchéance = mTope.AmjEchéance1
If mTope.optReprise = "R" Then
    chkComptaReprise = "1"
Else
    chkComptaReprise = "0"
End If
arrTFlux_Nb = 0

cmdControl

blncmdOk_Visible = True: blncmdSave_Visible = True
cmdOk.Visible = True
currentAction = "BasculeEuro"
cmdOk.Caption = currentAction
cmdContext.Caption = constcmdAbandonner

End Sub

Private Sub mnuPrêtModifierEchCompte_Click()
currentAction = "EchCompte"
cmdReset
fraPrêt_Load "Update"

fraPrêt.Enabled = True
fraNature.Enabled = False
fraRésultat.Enabled = False
Call lbl_Style(lblEchCompte, True)
fraCompta.Enabled = True
meEnabled_Container "fraCompta", False

txtEchCompte.Enabled = True
Call lstErr_Clear(lstErr, txtEchCompte, "> préciser le compte")
SSTab1.Tab = 2
blncmdOk_Visible = True
cmdOk.Caption = "OK_Compte"
cmdOk.FontSize = 6: cmdOk.FontName = "MS Serif"

cmdOk.Visible = True
End Sub

Private Sub mnucmdPrintList_TFlux_Click()
Dim Msg As String
If arrTFlux_Nb > 0 Then
    prtPrêt.recTFlux = recTFlux
    prtPrêt.CV1 = CV1
    ReDim prtPrêt.P_arrTFlux(arrTFlux_Nb)
    
    For I = 1 To fgEchéancier.Rows - 1
        fgEchéancier.Row = I
        fgEchéancier.Col = 11
        arrTFlux_Index = Val(fgEchéancier.Text)
        prtPrêt.P_arrTFlux(I) = arrTFlux(arrTFlux_Index) ' arrTOpe(I)
    Next I
Msg = Format$(1, "000000") & Format$(arrTFlux_Nb, "000000")

prtPrêtListEchéancier_Monitor Msg
End If

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

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case Is = "PRÊTSPP":  mNature = "PP_": Call BiaPgmAut_Init("PRÊTSPP", PrêtsAut): Me.Caption = "Prêts (service du PERSONNEL)"
    Case Is = "PRÊTSPC":  mNature = "PC_": Call BiaPgmAut_Init("PRÊTSPC", PrêtsAut): Me.Caption = "Prêts (service de la CAISSE)"
    Case Else: Unload Me
End Select
mnuPrêtSaisir.Enabled = PrêtsAut.Saisir

TFlux_Compta.param_Init mNature, cboNature
Form_Init " "
If UCase$(Trim(mId$(Msg, 13, 12))) = "BIA_EXPLOIT" Then
    mnuListEchéancier_Click
    mnucmdPrintList_TFlux_Click
    Unload Me
End If
End Sub

Public Sub Msg_Snd(ByVal X As String)
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

'-------------------------------------------------'
Private Sub txtAmjEchéance_control()
'-------------------------------------------------'

Dim X As String
X = Format$(txtAMJEchéance.Year, "0000") & Format$(txtAMJEchéance.Month, "00") & Format$(txtAMJEchéance.Day, "00")
If Not IsNumeric(X) Then
    Call lstErr_AddItem(lstErr, cmdContext, "? erreur date")
    DTPicker_Now txtAMJEchéance
Else
    wAmjEchéance = mId$(X, 1, 8)
End If

End Sub

'-------------------------------------------------'
Private Sub txtAmjEngagement_control()
'-------------------------------------------------'

Dim X As String
X = Format$(txtAmjEngagement.Year, "0000") & Format$(txtAmjEngagement.Month, "00") & Format$(txtAmjEngagement.Day, "00")
If Not IsNumeric(X) Then
    Call lstErr_AddItem(lstErr, cmdContext, "? erreur date")
    DTPicker_Now txtAmjEngagement
Else
    wAmjEngagement = mId$(X, 1, 8)
    txtAmjEchéance_Init
    
End If

End Sub

Private Sub cboNature_Click()
cbo_Value recTope.Nature, cboNature

End Sub


Private Sub cboNature_GotFocus()
lblNature.ForeColor = warnUsrColor
End Sub


Private Sub cboNature_LostFocus()
lblNature.ForeColor = lblUsr.ForeColor
If blnControl Then cmdControl
End Sub


Private Sub chkComptaReprise_Click()
If blnControl Then cmdControl

End Sub

Private Sub chkComptaReprise_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkComptaReprise

End Sub


Private Sub cmdContext_Click()

Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub


Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext
End Sub

Private Sub cmdOk_Click()
Dim mcmdOk_Caption As String

mcmdOk_Caption = cmdOk.Caption

If cmdOk.Caption = constàCompta Then
    cmdSave_àCompta
Else
    cmdControl
    If cmdOk.Caption = "FinAuto" Then lstErr.Clear
   If lstErr.ListCount <> 0 Then Exit Sub
    frmPrêt.Enabled = False
    Select Case cmdOk.Caption
        Case constàValider
            cmdPrint_Call constàValider
            recTope.Statut = "à"
            recTope.StatutPlus = "V "
            recTope.MajAMJ = DSys
            recTope.MajHMS = time_Hms
            recTope.MajUsr = usrId
        Case constValider
            If Not PrêtsAut.Xspécial And Trim(recTope.MajUsr) = Trim(usrId) Then
                Call MsgBox("Vous ne pouvez pas valider vos propres opérations.", vbCritical, "Garantie : Validation ")
                Call lstErr_AddItem(lstErr, cmdContext, "? validation interdite")
            Else
                fgEchéancier_AddNew
                recTope.Statut = " "
                recTope.StatutPlus = "  "
                recTope.valAMJ = DSys
                recTope.ValHMS = time_Hms
                recTope.ValUsr = usrId
                blnComptaAuto = PrêtsAut.Comptabiliser
            End If
        Case constRemboursementAnticipé
            fgEchéancier_RemboursementAnticipé
            fgEchéancier_Update
            recTope.Statut = "A"
            recTope.StatutPlus = "RA"
            recTope.valAMJ = DSys
            recTope.ValHMS = time_Hms
            recTope.ValUsr = usrId
        Case "OK_Compte"
            cmdPrint_Call "Modification du compte à préléver"
            recTope.MajAMJ = DSys
            recTope.MajHMS = time_Hms
            recTope.MajUsr = usrId
        Case "FinAuto"
            recTope.Method = constUpdate
            recTope.Statut = "F"
            recTope.StatutPlus = "in"
            
        Case "BasculeEuro"
            recTope.Method = constAddNew
            recTope.Statut = "à"
            recTope.StatutPlus = "E "
            recTope.ValUsr = ""
             V = srvTope_Update(recTope)
              If Not IsNull(V) Then
                Call lstErr_AddItem(lstErr, cmdContext, "? cmdOk : " & cmdOk.Caption)
            Else
                recTope.Method = constUpdate
                fgEchéancier_AddNew
                recTope.Statut = " "
                recTope.StatutPlus = "  "
                recTope.valAMJ = DSys
                recTope.ValHMS = time_Hms
                recTope.ValUsr = usrId
            End If
     Case Else
            Call lstErr_AddItem(lstErr, cmdContext, "? cmdOk : " & cmdOk.Caption)
    End Select

    If lstErr.ListCount = 0 Then
        cmdSave_Db
        If mcmdOk_Caption = constRemboursementAnticipé Then cmdRemboursementAnticipé_àCompta

    End If
    frmPrêt.Enabled = True
 '   AppActivate frmPrêt.Caption
End If

End Sub


Private Sub cmdPrint_Click()
Me.PopupMenu mnucmdPrint, vbPopupMenuLeftButton
End Sub

Private Sub cmdSave_Click()
cmdControl
lstErr.Clear
frmPrêt.Enabled = False
Select Case cmdSave.Caption
    Case constEnAttente
        recTope.MajAMJ = DSys
        recTope.MajHMS = time_Hms
        recTope.MajUsr = usrId
    Case constàModifier
        fgEchéancier_AddNew
        recTope.Statut = "à"
        recTope.StatutPlus = "? "
        recTope.valAMJ = DSys
        recTope.ValHMS = time_Hms
        recTope.ValUsr = constàModifier
    Case Else
        Call lstErr_AddItem(lstErr, cmdContext, "? cmdsave : " & cmdSave.Caption)
End Select

If lstErr.ListCount = 0 Then cmdSave_Db
frmPrêt.Enabled = True
End Sub
Public Sub cmdSave_Db()
If lstErr.ListCount = 0 Then
    blnControl = False
    
    V = srvTope_Update(recTope)
    
    If IsNull(V) Then
        If blnfgSelect_DisplayLine Then arrTOpe(arrTOpe_Index) = recTope: fgSelect_DisplayLine
        lastActiveControl_Name = ""
        cmdOk.Visible = False
        cmdSave.Visible = False
        Call lstErr_Clear(lstErr, cmdContext, "Mise à jour effectuée - Identification : " & recTope.IdRéférence)
        If blnComptaAuto Then mnuComptaDossier
        cmdContext_Quit
        SSTab1.Tab = 1
    Else
        Call lstErr_Clear(lstErr, cmdContext, V)
 ''''       cmdReset
    End If
End If
End Sub

Public Sub mnuComptaDossier()
arrTFlux_Index = 0
For I = 1 To arrTFlux_Nb
    If arrTFlux(I).Statut = " " And arrTFlux(I).AmjEchéanceTrt <= DSys Then
        arrTFlux_Index = arrTFlux_Index + 1
        arrTFlux(arrTFlux_Index) = arrTFlux(I)
    End If
Next I

If arrTFlux_Index = 0 Then
    cmdPrint_Call constValider
    cmdReset
    Exit Sub
End If

arrTFlux_Nb = arrTFlux_Index
cmdSave_àCompta
recTFlux = arrTFlux(1)
mnuLotàComptaValidation_Click

End Sub

Private Sub mnuListEchéancier_Click()
V = DTPicker_Control(txtAmjMax, wAmjEchéanceTrt)
If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, V): Exit Sub
cmdReset
mnuListEchéancier_Load

End Sub


Public Sub mnuListEchéancier_Load()
ReDim arrTFlux(1)

recTFlux_Init recTFlux
recTFlux.Method = "SnapLE"
recTFlux.AmjEchéanceTrt = "00000000"

arrTFlux(0) = recTFlux
arrTFlux(0).AmjEchéanceTrt = wAmjEchéanceTrt
arrTFlux(0).IdRéférence = 999999999
arrTFlux(0).IdSéquence = 32000

Call srvTFlux_Load(recTFlux, arrTFlux(0))
arrTFlux_Nb = srvTFlux.arrTFlux_Nb
ReDim arrTFlux(arrTFlux_Nb)
For I = 1 To arrTFlux_Nb
    arrTFlux(I) = srvTFlux.arrTFlux(I)
Next I
SSTab1.Tab = 3
fgEchéancier_Display " "

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
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
'fgSelect_FormatString = fgSelect.FormatString
'fgEchéancier_FormatString = fgEchéancier.FormatString
blnControl = False

End Sub

Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSelect.Row

If lRow > 0 Then
    fgSelect.Row = lRow
    For I = 0 To 11
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = 0 To 11
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.Col = 0
    End If
End If

End Sub

Public Sub fgEchéancier_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgEchéancier.Row

If lRow > 0 Then
    fgEchéancier.Row = lRow
    For I = 0 To 11
        fgEchéancier.Col = I: fgEchéancier.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgEchéancier.Row = mRow
    If fgEchéancier.Row > 0 Then
        lRow = fgEchéancier.Row
        lColor_Old = fgEchéancier.CellBackColor
        For I = 0 To 11
          fgEchéancier.Col = I: fgEchéancier.CellBackColor = lColor
        Next I
        fgEchéancier.Col = 0
    End If
End If

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub

Private Sub fraCompta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraNature_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraPrêt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraRésultat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub mnuAbandonner_Click()
cmdContext_Quit
End Sub

Private Sub mnucmdPrintList_TOpe_Click()
Dim Msg As String
If arrTOpe_NB > 0 Then
    prtPrêt.recTope = recTope
    prtPrêt.CV1 = CV1
    ReDim prtPrêt.P_arrTOpe(arrTOpe_NB)
    For I = 1 To arrTOpe_NB
        prtPrêt.P_arrTOpe(I) = arrTOpe(I)
    Next I
    Msg = Format$(1, "000000") & Format$(arrTOpe_NB, "000000")
    prtPrêtList_Monitor Msg
End If

End Sub

Private Sub mnucmdPrintPrêt_Click()
cmdPrint_Call "Prêt"
End Sub

Private Sub mnucmdPrintTableauAmortissement_Click()
cmdPrint_Call "Tableau"
End Sub


Private Sub mnuComptaEchéancier_Click()
wAmjEchéanceTrt = DSys
mnuComptaEchéancier_Load
End Sub

Private Sub mnuComptaEchéancier_Plus_Click()
V = DTPicker_Control(txtAmjMax, wAmjEchéanceTrt)
If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, V): Exit Sub

If mId$(wAmjEchéanceTrt, 1, 6) <> mId$(DSys, 1, 6) Then Call lstErr_AddItem(lstErr, cmdContext, "? date échéance > mois en cours)"): Exit Sub

X = MsgBox("Confirmez-vous la date du : " & dateImp(wAmjEchéanceTrt) & " ?", vbYesNo + vbQuestion + vbDefaultButton2, "frmPrêt : Comptabilisation anticipée des échéances")
If X = vbYes Then mnuComptaEchéancier_Load

End Sub


Private Sub mnuComptaLotsàValider_Click()
ReDim arrTFlux(1)

recTFlux_Init recTFlux
recTFlux.Method = "SnapLotàC"
recTFlux.Statut = "à"

arrTFlux(0) = recTFlux
arrTFlux(0).CptMvtLot = "99999999"
Call srvTFlux_Load(recTFlux, arrTFlux(0))
arrTFlux_Nb = srvTFlux.arrTFlux_Nb
ReDim arrTFlux(arrTFlux_Nb)
For I = 1 To arrTFlux_Nb
    arrTFlux(I) = srvTFlux.arrTFlux(I)
    arrTFlux(I).CodeOpération = "$Lot"
    arrTFlux(I).Capital = 0
    arrTFlux(I).Intérêts = 0
Next I
SSTab1.Tab = 3
fgEchéancier_Display " "

If arrTFlux_Nb = 0 Then
    MsgBox "frmPrêt mnuComptaLotsàValider : PAS DE LOTS à TRAITER"
Else
    currentAction = constàCompta_Valider
End If

End Sub

Private Sub mnuListàValider_Click()
currentAction = "mnuListàValider"
fgSelect_Load
End Sub

Private Sub mnuListPrêts_Click()
currentAction = "mnuListPrêts"
fgSelect_Load
End Sub


Private Sub mnuLotàComptaAnnulation_Click()
Call lstErr_Clear(lstErr, cmdContext, "!! Suppression de l'échéancier")
recTFlux.Method = constàCompta_Annuler
recTFlux.Statut = "à"
Call srvTFlux_Update(recTFlux)
fgEchéancier.Clear: fgEchéancier.Rows = 1
End Sub

Private Sub mnuLotàComptaPrint_Click()
TFlux_Compta.LotàCompta_Demande recTFlux.CptMvtLot

End Sub

Private Sub mnuLotàComptaValidation_Click()
Dim X As String

frmPrêt.Enabled = False
If blnComptaAuto Then
    X = vbYes
Else
    X = MsgBox("Cette action est irréversible. Confirmez-vous votre demande ?", vbYesNo + vbQuestion + vbDefaultButton2, "tflux : Validation définitive du lot")
End If
If X = vbYes Then

    recTFlux.Method = constàCompta_Valider
    recTFlux.CptMvtUsr = usrId
    recTFlux.CptMvtAMJ = DSys
    recTFlux.CptMvtHMS = time_Hms
    mTflux = recTFlux
    srvTFlux_Update recTFlux
    TFlux_Compta.LotàCompta_Valider mTflux
    cmdReset
End If

frmPrêt.Enabled = True
AppActivate frmPrêt.Caption

End Sub

Private Sub mnuLotComptabilisé_Annuler_Click()
X = InputBox("Indiquer le numéro du lot à annuler :")
recTFlux_Init recTFlux
recTFlux.Statut = "C"
recTFlux.Method = "Compta_Ann"
recTFlux.CptMvtLot = CLng(Val(X))
If recTFlux.CptMvtLot = 0 Then
    Call lstErr_Clear(lstErr, cmdContext, " ! N° lot = 0 ")
Else
    srvTFlux_Update recTFlux
    Call lstErr_Clear(lstErr, cmdContext, " annulation comptabilisation du lot " & X)
End If

End Sub

Private Sub mnuPrêtDisplay_Click()
fraPrêt_Load " "

Dim I As Integer, blnTest As Boolean
If PrêtsAut.Xspécial And mTope.Statut = " " Then
    blnTest = True
    For I = 1 To arrTFlux_Nb
            If arrTFlux(I).Statut = " " Then blnTest = False
    Next I
    If blnTest Then
        blncmdOk_Visible = True: blncmdSave_Visible = False
        cmdOk.Caption = "FinAuto"
        cmdOk.Visible = True
        cmdContext.Caption = constcmdAbandonner
    End If
End If

End Sub

Private Sub mnuPrêtEffacer_Click()
fraPrêt_Load "Delete"

End Sub

Private Sub mnuPrêtModifier_Click()
mnuPrêtSaisir_Click
fraPrêt_Load "Update"
blncmdOk_Visible = True: blncmdSave_Visible = True
End Sub

Private Sub mnuPrêtRemboursementAnticipé_Click()
Dim I As Integer

'fraPrêt_Load constRemboursementAnticipé
currentAction = constRemboursementAnticipé
'mnuPrêtSaisir_Click
fraPrêt.Enabled = False
fraPrêt_Load "Update"
blncmdOk_Visible = True: blncmdSave_Visible = False
cmdOk.Caption = constRemboursementAnticipé
cmdOk.Visible = True
cmdContext.Caption = constcmdAbandonner
wAmjRemboursementAnticipé = DSys
txtAMJ.Visible = True
Call DTPicker_Set(txtAMJ, DSys)
minAmjRemboursementAnticipé = 99991231
For I = 1 To arrTFlux_Nb
    If arrTFlux(I).CodeOpération = "PR02" Then
        If arrTFlux(I).Statut <> " " Then minAmjRemboursementAnticipé = arrTFlux(I).AmjFin
    End If
Next I

End Sub

Private Sub mnuPrêtSaisir_Click()
If PrêtsAut.Saisir Then
    SSTab1.Tab = 2
    currentAction = constSaisie
    fraSelect.Enabled = False
    fgSelect.Enabled = False
    fgEchéancier.Enabled = False
    fgEchéancier.Clear: fgEchéancier.Rows = 1
    cmdReset
    fraPrêt.Enabled = True
    blncmdOk_Visible = True: blncmdSave_Visible = True
    blnAmjEchéance = False
    txtCapital.SetFocus
    cmdContext.Caption = constcmdAbandonner
End If

End Sub

Private Sub mnuPrêtValider_Click()
currentAction = constValider
mnuPrêtSaisir_Click
fraPrêt.Enabled = False
fraPrêt_Load "Update"
blncmdOk_Visible = True: blncmdSave_Visible = True
cmdOk.Visible = True
currentAction = constValider
cmdContext.Caption = constcmdAbandonner
End Sub

Private Sub mnuQuitter_Click()
Unload Me
End Sub

Private Sub optAnnuel_Click()
If blnControl Then cmdControl
End Sub

Private Sub optAnnuel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optAnnuel
End Sub


Private Sub optEchéanceAnniversaire_Click()
If blnControl Then cmdControl
End Sub

Private Sub optEchéanceAnniversaire_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEchéanceAnniversaire
End Sub


Private Sub optEchéanceFinDeMois_Click()
If blnControl Then cmdControl
End Sub

Private Sub optEchéanceFinDeMois_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEchéanceFinDeMois

End Sub


Private Sub optMensuel_Click()
If blnControl Then cmdControl
End Sub


Private Sub optMensuel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optMensuel

End Sub


Private Sub optSemestriel_Click()
If blnControl Then cmdControl
End Sub

Private Sub optSemestriel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSemestriel
End Sub


Private Sub optTrimestriel_Click()
If blnControl Then cmdControl
End Sub


Private Sub optTrimestriel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optTrimestriel
End Sub


Private Sub txtAmjEchéance_Change()
blnAmjEchéance = True
txtAmjEchéance_control

End Sub

Private Sub txtAmjEchéance_GotFocus()
DTPicker_GotFocus txtAMJEchéance
End Sub

Private Sub txtAMJEchéance_LostFocus()
DTPicker_LostFocus txtAMJEchéance
txtAmjEchéance_control
If blnControl Then cmdControl

End Sub

Private Sub txtAmjEchéance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DTPicker_GotFocus txtAMJEchéance
End Sub

Private Sub txtAmjEngagement_Change()
txtAmjEngagement_control
End Sub

Private Sub txtAmjEngagement_GotFocus()
DTPicker_GotFocus txtAmjEngagement

End Sub

Private Sub txtAmjEngagement_LostFocus()
DTPicker_LostFocus txtAmjEngagement
txtAmjEngagement_control
If blnControl Then cmdControl

End Sub

Private Sub txtAmjEngagement_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DTPicker_GotFocus txtAmjEngagement

End Sub

Public Sub cmdControl()
Dim X As String, wMensualité As Currency, wAmj As String, wTaux As Double
Dim blnControlBiatyp As Boolean

If Not frmPrêt.Enabled Then Exit Sub
frmPrêt.Enabled = False

cmdOk.Visible = False
blnControl = False

lstErr.Clear
lstErr.Height = 200

recTope = mTope
If currentAction = constSaisie Or currentAction = constSaisie Then
    blnControlBiatyp = True
Else
    blnControlBiatyp = False
End If

If currentAction = constRemboursementAnticipé Then
    Call DTPicker_Control(txtAMJ, wAmjRemboursementAnticipé)
    If wAmjRemboursementAnticipé < minAmjRemboursementAnticipé Then Call lstErr_AddItem(lstErr, cmdContext, "? Date remboursement impossible")
End If

X = num_Control(txtIdRéférence, valX, 10, 0)
recTope.IdRéférence = valX

CV1.DeviseIso = Trim(txtDevise)
V = CV_Attribut(CV1)
If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, V)

X = num_Control(txtCapital, valX, 13, CV1.maxD)
recTope.Capital = valX
If recTope.Capital <= 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le capital emprunté")

X = num_Control(txtTaux, valX, 9, 5)
recTope.TauxMarge = valX
If recTope.TauxMarge <= 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le taux")
If recTope.TauxMarge > 25 Then Call lstErr_AddItem(lstErr, cmdContext, "? taux > 25 %")

X = num_Control(txtPériodeNb, valX, 4, 0)
recTope.PériodeNb = valX
If recTope.PériodeNb <= 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le nombre de période")
If recTope.PériodeNb > 360 Then Call lstErr_AddItem(lstErr, cmdContext, "? nombre de période > 360")

If optMensuel Then recTope.Périodicité = "M": fctPériodicité = "MoisAdd"
If optTrimestriel Then recTope.Périodicité = "T": fctPériodicité = "TrimestreAdd"
If optSemestriel Then recTope.Périodicité = "S": fctPériodicité = "SemestreAdd"
If optAnnuel Then recTope.Périodicité = "A": fctPériodicité = "AnAdd"

X = num_Control(txtFrais, valX, 13, CV1.maxD)
recTope.Frais = valX

If lstErr.ListCount > 0 Then GoTo ExitSub
txtAmjEchéance_Init

recTope.Application = paramTFlux_Service
recTope.Devise = CV1.DeviseIso
recTope.IPA = "E"
recTope.NbjBase = "0"
Call cbo_Value(recTope.Nature, cboNature)
V = TFlux_Compta.param_Nature(recTope.Nature)
If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, V)
X = Trim(txtRéférenceInterne)
recTope.RéférenceInterne = X

cmdSave.Visible = False
If X = "" Then
    Call lstErr_AddItem(lstErr, cmdContext, "? préciser la référence interne")
Else
    cmdSave.Visible = blncmdSave_Visible
End If
recTope.RéférenceExterne = Trim(txtRéférenceExterne)

txtAmjEngagement_control: recTope.AmjDébut = wAmjEngagement
txtAmjEchéance_control: recTope.AmjEchéance1 = wAmjEchéance
recTope.AmjFin = dateElp(fctPériodicité, recTope.PériodeNb - 1, recTope.AmjEchéance1)

If currentAction = constSaisie Or currentAction = constValider Then
    If recTope.AmjDébut < paramAmjEngagementMin And chkComptaReprise = "0" Then Call lstErr_AddItem(lstErr, cmdContext, "? Reprise : cocher la case")
    If recTope.AmjDébut > paramAmjEngagementMax Then Call lstErr_AddItem(lstErr, cmdContext, "? date du prêt > 7 jours")
End If

paramAmjEchéanceMin = dateElp(fctPériodicité, 1, recTope.AmjDébut)
paramAmjEchéanceMax = dateElp(fctPériodicité, 3, recTope.AmjDébut)
If recTope.AmjEchéance1 < paramAmjEchéanceMin Then Call lstErr_AddItem(lstErr, cmdContext, "? 1 ère échéance <= 1 période")
If recTope.AmjEchéance1 > paramAmjEchéanceMax Then Call lstErr_AddItem(lstErr, cmdContext, "? 1 ère échéance >= 3 période")

If optEchéanceFinDeMois Then
    recTope.AmjEchéanceS = "M"
Else
    recTope.AmjEchéanceS = "A"
End If

If recTope.AmjEchéanceS = "M" Then
    recTope.AmjFin = dateFinDeMois(recTope.AmjFin)
    wAmj = dateFinDeMois(recTope.AmjEchéance1)
    If recTope.AmjEchéance1 <> wAmj Then Call lstErr_AddItem(lstErr, cmdContext, recTope.AmjEchéance1 & " ? n'est pas une fin de mois")
End If

Call fctTOpe_Mensualité(recTope, CV1, wMensualité, wTaux, recTope.TauxActuariel, recTope.TEG)
recTope.Mensualité = wMensualité
txtMensualité = Format$(recTope.Mensualité, "### ### ### ##0.00")
txtTauxActuariel = Format$(recTope.TauxActuariel, "##0.00000")
Call TEG_Calc(recTope.Capital, recTope.Frais, recTope.Mensualité, recTope.PériodeNb, recTope.Périodicité, recTope.TauxMarge, recTope.TEG)
recTope.TEG = Round(recTope.TEG, 2)
txtTEG = Format$(recTope.TEG, "##0.00")

libEchCompte = ""
If Trim(txtEchCompte) = "" Then
    Call lstErr_AddItem(lstErr, cmdContext, "? préciser le compte à préléver")
Else
    X = num_Control(txtEchCompte, valX, 11, 0)
    cmdControl_Compte valX
    recTope.EchéanceCompte = valX
    V = Compte_Load(recTope.EchéanceCompte)
    libEchCompte = Trim(recCompte.Intitulé) & " / " & Trim(recCompte.Intitulé2)
    txtEchCompte = Compte_Display(recTope.EchéanceCompte)
    If IsNull(V) And blnControlBiatyp Then
        If recTope.EchéanceCompte <> mEChéanceCompte Then
            If mId$(recTope.EchéanceCompte, 6, 3) <> paramTFlux_BiatypEchéance Then
                X = MsgBox("Le type de compte attendu est : " & paramTFlux_BiatypEchéance & Chr$(13) & " confirmez-vous ce compte?", vbYesNo + vbQuestion + vbDefaultButton2, "Prêt : Compte à prélever ")
                If X = vbYes Then mEChéanceCompte = recTope.EchéanceCompte
            End If
        End If
    End If
    
End If


libEngCompte = ""
If Trim(txtEngCompte) = "" Then
    Call lstErr_AddItem(lstErr, cmdContext, "? préciser le compte d'engagement")
Else
    X = num_Control(txtEngCompte, valX, 11, 0)
    recTope.EngagementCompte = valX
    V = Compte_Load(recTope.EngagementCompte)
    libEngCompte = Trim(recCompte.Intitulé) & " / " & Trim(recCompte.Intitulé2)
    txtEngCompte = Compte_Display(recTope.EngagementCompte)
    If IsNull(V) And blnControlBiatyp Then
        If recTope.EngagementCompte <> mEngagementCompte Then
            If mId$(recTope.EngagementCompte, 6, 3) <> paramTFlux_BiatypEngagement Then
                X = MsgBox("Le type de compte attendu est : " & paramTFlux_BiatypEngagement & Chr$(13) & " confirmez-vous ce compte?", vbYesNo + vbQuestion + vbDefaultButton2, "Prêt : Compte de prêt ")
                If X = vbYes Then mEngagementCompte = recTope.EngagementCompte
            End If
        End If
    End If
End If

libCorrCompte = ""
If Trim(txtCorrCompte) = "" Then
    Call lstErr_AddItem(lstErr, cmdContext, "? préciser le compte correspondant")
Else
    X = num_Control(txtCorrCompte, valX, 11, 0)
    recTope.EngagementCorrCompte = valX
    V = Compte_Load(recTope.EngagementCorrCompte)
    libCorrCompte = Trim(recCompte.Intitulé) & " / " & Trim(recCompte.Intitulé2)
    txtCorrCompte = Compte_Display(recTope.EngagementCorrCompte)
    If IsNull(V) And blnControlBiatyp Then
        If recTope.EngagementCorrCompte <> mEngagementCorrCompte Then
            If mId$(recTope.EngagementCorrCompte, 6, 3) <> paramTFlux_BiatypEngagementCorr Then
                X = MsgBox("Le type de compte attendu est : " & paramTFlux_BiatypEngagementCorr & Chr$(13) & " confirmez-vous ce compte?", vbYesNo + vbQuestion + vbDefaultButton2, "Prêt : Compte de déblocage des fonds ")
                If X = vbYes Then mEngagementCorrCompte = recTope.EngagementCorrCompte
            End If
        End If
    End If
    
End If


If chkComptaReprise = "1" Then
    recTope.optReprise = "R"
Else
    recTope.optReprise = " "
End If

If recTope.Statut = "à" Then
    Call fgEchéancier_Gen(wTaux)
    fgEchéancier_Display "T"
End If

If currentAction = constValider Then
    V = fctTOpe_Compare(recTope, mTope)
    If Not IsNull(V) Then
        Call MsgBox("L'enregistrement après contrôle est différent de l'enregistrement lu :" & Chr$(13) & V, vbCritical, "frmPrêt : cmdControl")
        Call lstErr_AddItem(lstErr, cmdContext, "? Erreur Contrôle validation")
    End If
End If

If lstErr.ListCount = 0 Then
    cmdOk.Visible = blncmdOk_Visible
End If

ExitSub:

frmPrêt.Enabled = True
If cmdOk.Visible Then cmdOk.SetFocus 'mdOk.Visible = False: cmdOk.Visible = True
    
blnControl = True

End Sub

Public Sub fgEchéancier_Display(Fct As String)
Dim I As Integer

fgEchéancier.Visible = True
fgEchéancier.Clear
totalCapital = 0: totalIntérêts = 0

fgEchéancier.Rows = 1
fgEchéancier.FormatString = fgEchéancier_FormatString
fgEchéancier.Enabled = True
For arrTFlux_Index = 1 To arrTFlux_Nb
    recTFlux = arrTFlux(arrTFlux_Index)
    fgEchéancier.Rows = fgEchéancier.Rows + 1
    fgEchéancier.Row = fgEchéancier.Rows - 1
    fgEchéancier_DisplayLine
Next arrTFlux_Index

fgEchéancier_K = fgEchéancier.Cols
 
If Fct = "T" And fgEchéancier.Rows > 1 Then
    fgEchéancier.TextArray(3 + fgEchéancier_K) = Format(totalCapital, "#### ### ###.00 ")
    fgEchéancier.TextArray(4 + fgEchéancier_K) = Format(totalIntérêts, "#### ### ##0.00 ")
End If

'if totalcapital
''If fgEchéancier.Rows > 1 Then fgEchéancier_Sort

End Sub

Public Sub fgEchéancier_DisplayLine()
Dim K2 As Integer

fgEchéancier_K = (fgEchéancier.Row) * fgEchéancier.Cols
If recTFlux.CodeOpération = "$Lot" Then
     fgEchéancier.TextArray(0 + fgEchéancier_K) = " Lot : " & Format(recTFlux.CptMvtLot, "### ### ")
    fgEchéancier.TextArray(1 + fgEchéancier_K) = " Lot à comptabiliser"
Else
 
    fgEchéancier.TextArray(0 + fgEchéancier_K) = dateImp(recTFlux.AmjEchéanceTrt)
    fgEchéancier.TextArray(1 + fgEchéancier_K) = recTFlux.CodeOpération
End If

fgEchéancier.TextArray(2 + fgEchéancier_K) = Format(recTFlux.Capital + recTFlux.Intérêts, "#### ### ###.00 ")
If recTFlux.Capital <> 0 Then fgEchéancier.TextArray(3 + fgEchéancier_K) = Format(recTFlux.Capital, "#### ### ##0.00 ")
If recTFlux.Intérêts <> 0 Then fgEchéancier.TextArray(4 + fgEchéancier_K) = Format(recTFlux.Intérêts, "#### ### ##0.00 ")
fgEchéancier.TextArray(5 + fgEchéancier_K) = Format(recTFlux.Taux, "#0.00000 ") & recTFlux.TauxProvisoire
fgEchéancier.TextArray(6 + fgEchéancier_K) = "du " & dateImp(recTFlux.AmjDébut) & " au " & dateImp(recTFlux.AmjFin) & "   (" & recTFlux.Nbj & "j)"
fgEchéancier.TextArray(7 + fgEchéancier_K) = recTFlux.CptMvtUsr & " " & dateImp(recTFlux.CptMvtAMJ) & " " & timeImp(recTFlux.CptMvtHMS)
fgEchéancier.TextArray(8 + fgEchéancier_K) = Format(recTFlux.CptMvtLot, "### ### ") ''''& Format(recTFlux.CptMvtPièce, "### ### ") & Format(recTFlux.CptMvtLigne, "### ### ")
fgEchéancier.TextArray(9 + fgEchéancier_K) = Format(recTFlux.IdRéférence, "### ##0 ") & "_" & Format(recTFlux.IdSéquence, "### ### ")
fgEchéancier.TextArray(10 + fgEchéancier_K) = ""
fgEchéancier.TextArray(10 + fgEchéancier_K) = recStatut_Libellé(recTFlux.Statut & recTFlux.StatutPlus)
fgEchéancier.TextArray(11 + fgEchéancier_K) = arrTFlux_Index

If recTFlux.CodeOpération = "PR01" Or recTFlux.CodeOpération = "PR05" Then
    fgEchéancier.Col = 2: fgEchéancier.CellForeColor = errUsr.ForeColor
    fgEchéancier.Col = 3: fgEchéancier.CellForeColor = errUsr.ForeColor
    fgEchéancier.Col = 4: fgEchéancier.CellForeColor = errUsr.ForeColor
Else
    totalCapital = totalCapital + recTFlux.Capital
    totalIntérêts = totalIntérêts + recTFlux.Intérêts
End If

End Sub

Private Sub txtAmjMax_GotFocus()
DTPicker_GotFocus txtAmjMax

End Sub

Private Sub txtAmjMax_LostFocus()
DTPicker_LostFocus txtAmjMax
End Sub

Private Sub txtAmjMin_GotFocus()
DTPicker_GotFocus txtAmjMin

End Sub

Private Sub txtAmjMin_LostFocus()
DTPicker_LostFocus txtAmjMin

End Sub

Private Sub txtCapital_GotFocus()
txt_GotFocus txtCapital

End Sub


Private Sub txtCapital_KeyPress(KeyAscii As Integer)
If CV1.maxD = 0 Then
    Call num_KeyAscii(KeyAscii)
Else
    Call num_KeyAsciiD(KeyAscii, txtCapital)
End If

End Sub


Private Sub txtCapital_LostFocus()
txt_LostFocus txtCapital
If blnControl Then cmdControl

End Sub

Private Sub txtCorrCompte_GotFocus()
txt_GotFocus txtCorrCompte

End Sub


Private Sub txtCorrCompte_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtCorrCompte)

End Sub


Private Sub txtCorrCompte_LostFocus()
txt_LostFocus txtCorrCompte
If blnControl Then cmdControl

End Sub

Private Sub txtDevise_GotFocus()
txt_GotFocus txtDevise

End Sub


Private Sub txtDevise_LostFocus()
txt_LostFocus txtDevise
If blnControl Then cmdControl

End Sub


Private Sub txtEchCompte_GotFocus()
txt_GotFocus txtEchCompte

End Sub


Private Sub txtEchCompte_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtEchCompte)

End Sub


Private Sub txtEchCompte_LostFocus()
txt_LostFocus txtEchCompte
If blnControl Then cmdControl

End Sub

Private Sub txtEngCompte_GotFocus()
txt_GotFocus txtEngCompte

End Sub


Private Sub txtEngCompte_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtEngCompte)

End Sub


Private Sub txtEngCompte_LostFocus()
txt_LostFocus txtEngCompte
If blnControl Then cmdControl

End Sub

Private Sub txtFrais_GotFocus()
txt_GotFocus txtFrais

End Sub


Private Sub txtFrais_KeyPress(KeyAscii As Integer)
If CV1.maxD = 0 Then
    Call num_KeyAscii(KeyAscii)
Else
    Call num_KeyAsciiD(KeyAscii, txtFrais)
End If

End Sub


Private Sub txtFrais_LostFocus()
txt_LostFocus txtFrais
If blnControl Then cmdControl

End Sub

Private Sub txtIdRéférence_LostFocus()
If blnControl Then cmdControl

End Sub


Private Sub txtMensualité_GotFocus()
txt_GotFocus txtMensualité

End Sub


Private Sub txtMensualité_KeyPress(KeyAscii As Integer)
If CV1.maxD = 0 Then
    Call num_KeyAscii(KeyAscii)
Else
    Call num_KeyAsciiD(KeyAscii, txtMensualité)
End If

End Sub


Private Sub txtMensualité_LostFocus()
txt_LostFocus txtMensualité
If blnControl Then cmdControl

End Sub

Private Sub txtPériodeNb_GotFocus()
txt_GotFocus txtPériodeNb

End Sub


Private Sub txtPériodeNb_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtPériodeNb)
End Sub


Private Sub txtPériodeNb_LostFocus()
txt_LostFocus txtPériodeNb
If blnControl Then cmdControl

End Sub

Private Sub txtRéférenceExterne_GotFocus()
txt_GotFocus txtRéférenceExterne

End Sub


Private Sub txtRéférenceExterne_LostFocus()
txt_LostFocus txtRéférenceExterne
If blnControl Then cmdControl

End Sub


Private Sub txtRéférenceInterne_GotFocus()
txt_GotFocus txtRéférenceInterne

End Sub


Private Sub txtRéférenceInterne_LostFocus()
txt_LostFocus txtRéférenceInterne
If blnControl Then cmdControl

End Sub


Private Sub txtSelect_GotFocus()
txt_GotFocus txtSelect
End Sub


Private Sub txtSelect_LostFocus()
txt_LostFocus txtSelect
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

Private Sub txtTauxActuariel_GotFocus()
txt_GotFocus txtTauxActuariel

End Sub


Private Sub txtTauxActuariel_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtTauxActuariel)

End Sub


Private Sub txtTauxActuariel_LostFocus()
txt_LostFocus txtTauxActuariel
If blnControl Then cmdControl

End Sub

Private Sub txtTEG_GotFocus()
txt_GotFocus txtTEG

End Sub


Private Sub txtTEG_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtTEG)

End Sub


Private Sub txtTEG_LostFocus()
txt_LostFocus txtTEG
If blnControl Then cmdControl

End Sub



Public Sub fgEchéancier_Gen(mTaux As Double)
Dim V As Variant, MsgErr As String, I As Integer, wNbj As Long, wIntérêts As Currency
Dim wAmjDébut  As String, wAmjFin As String

Dim totalAmortissement As Currency
On Error GoTo Error_Handle

ReDim arrTFlux(recTope.PériodeNb + 5)
xTOpe = recTope
srvTFlux.recTFlux_Init recTFlux

fgEchéancier.Clear: fgEchéancier.Rows = 1

With recTFlux                                   ' Engagement
    .IdSéquence = 1
    .CodeOpération = "PR01"
    .Capital = recTope.Capital
    .Taux = recTope.TauxMarge
    .AmjEchéanceTrt = recTope.AmjDébut
    .AmjDébut = recTope.AmjDébut
    .AmjFin = recTope.AmjFin
    .AmjOpération = recTope.AmjDébut
    .AmjValeur = recTope.AmjDébut
End With
arrTFlux(recTFlux.IdSéquence) = recTFlux

If recTope.Frais <> 0 Then
    With recTFlux                                   ' Frais
        .IdSéquence = recTFlux.IdSéquence + 1
        .CodeOpération = "PR03"
        .Capital = 0
        .Nbj = 0
        .Intérêts = recTope.Frais
        .AmjEchéanceTrt = recTope.AmjDébut
        .AmjDébut = recTope.AmjDébut
        .AmjFin = recTope.AmjDébut
        .AmjOpération = recTope.AmjDébut
        .AmjValeur = recTope.AmjDébut
    End With
    arrTFlux(recTFlux.IdSéquence) = recTFlux

End If

wAmjFin = recTope.AmjEchéance1
V = fctTOpe_AmjFinPrécédente(xTOpe, wAmjFin)
xTOpe.AmjFin = wAmjFin
If Not IsNull(V) Then MsgErr = "? calcul date début de la première période": Error 9999
V = fctTOpe_Intérêts(xTOpe, CV1, wIntérêts, wNbj)
If Not IsNull(V) Then MsgErr = "? calcul intérêts décalés": Error 9999

If wIntérêts <> 0 Then
    With recTFlux                                   ' intérêts décalés
        .IdSéquence = recTFlux.IdSéquence + 1
        .CodeOpération = "PR04"
        .Capital = 0
        .Intérêts = wIntérêts
        .Nbj = wNbj
        .AmjEchéanceTrt = wAmjFin
        .AmjDébut = recTope.AmjDébut
        .AmjFin = wAmjFin
        .AmjOpération = wAmjFin
        .AmjValeur = wAmjFin
    End With
    arrTFlux(recTFlux.IdSéquence) = recTFlux
End If

recTFlux.CodeOpération = "PR02"                     ' Echéance
totalAmortissement = recTope.Capital
For I = 1 To recTope.PériodeNb
    V = fctTOpe_PériodeSuivante(xTOpe, wAmjDébut, wAmjFin)
    If Not IsNull(V) Then MsgErr = "? calcul échéance " & I: Error 9999
    
    With recTFlux                                   ' intérêts décalés
        .IdSéquence = recTFlux.IdSéquence + 1
        .Intérêts = Round(totalAmortissement * mTaux, CV1.maxD)
        .Capital = xTOpe.Mensualité - recTFlux.Intérêts
        .Nbj = 30
        .AmjEchéanceTrt = wAmjFin
        .AmjDébut = wAmjDébut
        .AmjFin = wAmjFin
        .AmjOpération = wAmjFin
        .AmjValeur = wAmjFin
    End With
    arrTFlux(recTFlux.IdSéquence) = recTFlux
    totalAmortissement = totalAmortissement - recTFlux.Capital
Next I

If totalAmortissement <> 0 Then
    recTFlux.Capital = recTFlux.Capital + totalAmortissement
    recTFlux.Intérêts = xTOpe.Mensualité - recTFlux.Capital
    arrTFlux(recTFlux.IdSéquence) = recTFlux
End If
arrTFlux_Nb = recTFlux.IdSéquence
Exit Sub
'---------------------------------------------------------
Error_Handle:
'---------------------------------------------------------

Call MsgBox(MsgErr, vbCritical, "frmPrêts.fgEchéancier_Gen")


End Sub

Public Function Compte_Load(mCompteNuméro As String)
Compte_Load = Null
recCompteInit recCompte
recCompte.Société = SocId$
recCompte.Agence = SocAgence$
recCompte.Devise = CV1.DeviseN
recCompte.Numéro = mCompteNuméro
recCompte.BiaTyp = "000"
recCompte.BiaNum = "00"
recCompte.Method = "SeekL1"
If Not IsNull(srvCompteFind(recCompte)) Then Call lstErr_AddItem(lstErr, lstErr, "? compte inconnu : " & mCompteNuméro): Compte_Load = "?": Exit Function

If recCompte.Situation <> " " Then
    Select Case recCompte.Situation
        Case "B": Call lstErr_AddItem(lstErr, lstErr, " ? Compte bloqué : " & mCompteNuméro): Compte_Load = "?"
        Case "A": Call lstErr_AddItem(lstErr, lstErr, " ? Compte annulé : " & mCompteNuméro): Compte_Load = "?"
        Case "E": '2001.08.06 JPL   ne pas interdire
        Case Else: Call lstErr_AddItem(lstErr, lstErr, " ? Situation du compte : " & mCompteNuméro): Compte_Load = "?"
    End Select
End If

End Function

Public Sub cmdPrint_Call(Fct As String)
Dim Msg As String
If arrTFlux_Nb > 0 Then
    prtPrêt.recTope = recTope
    prtPrêt.CV1 = CV1
    ReDim prtPrêt.P_arrTFlux(arrTFlux_Nb)
    For I = 1 To arrTFlux_Nb
        prtPrêt.P_arrTFlux(I) = arrTFlux(I)
    Next I
    Msg = Format$(1, "000000") & Format$(arrTFlux_Nb, "000000") & Fct
    prtPrêt_Monitor Msg
End If

End Sub

Public Sub fraPrêt_Load(Fct As String)
'2000-01-04 cmdReset
fgSelect_RowClick = 0
Call fgSelect_Color(fgSelect_RowDisplay, vbCyan, fgSelect_ColorClick) 'txtUsr.BackColor)
xTOpe.Method = "SeekP0"
V = srvTOpe_Monitor(xTOpe)
If IsNull(V) Then
    blnControl = False: blnAmjEchéance = True
    SSTab1.Tab = 2
    mTope = xTOpe
    mTope.Method = Fct
    txtIdRéférence = Format$(mTope.IdRéférence, "#### ### ##0")
    cbo_Scan mTope.Nature, cboNature
    txtDevise = mTope.Devise
    txtCapital = Format$(mTope.Capital, "### ### ### ##0.00")
    txtTaux = Format$(mTope.TauxMarge, "#0.00000")
    txtPériodeNb = Format$(mTope.PériodeNb, "###0")
    Select Case mTope.Périodicité
        Case "M": optMensuel = True: fctPériodicité = "MoisAdd"
        Case "T": optTrimestriel = True: fctPériodicité = "TrimestreAdd"
        Case "S": optSemestriel = True: fctPériodicité = "SemestreAdd"
        Case "A": optAnnuel = True: fctPériodicité = "AnAdd"
        Case Else: optMensuel = True: fctPériodicité = "MoisAdd"
   End Select
If optAnnuel Then recTope.Périodicité = "A": fctPériodicité = "AnAdd"
  
    txtFrais = Format$(mTope.Frais, "### ### ### ##0.00")
    txtMensualité = Format$(mTope.Mensualité, "### ### ### ##0.00")
    txtTEG = Format$(mTope.TEG, "##0.00000")
    txtTauxActuariel = Format$(mTope.TauxActuariel, "##0.00000")
    txtEchCompte = Compte_Display(mTope.EchéanceCompte)
    txtEngCompte = Compte_Display(mTope.EngagementCompte)
    txtCorrCompte = Compte_Display(mTope.EngagementCorrCompte)
    txtRéférenceInterne = Trim(mTope.RéférenceInterne)
    txtRéférenceExterne = Trim(mTope.RéférenceExterne)
    Call DTPicker_Set(txtAmjEngagement, mTope.AmjDébut): wAmjEngagement = mTope.AmjDébut
    Call DTPicker_Set(txtAMJEchéance, mTope.AmjEchéance1): wAmjEchéance = mTope.AmjEchéance1
    If mTope.AmjEchéanceS = "M" Then
        optEchéanceFinDeMois = True
    Else
        optEchéanceAnniversaire = True
    End If
    
    If mTope.optReprise = "R" Then
        chkComptaReprise = "1"
    Else
        chkComptaReprise = "0"
    End If
    
    cmdControl
    
    If mTope.Statut = "à" Then
        Select Case mTope.StatutPlus
            Case Is = "V ": cmdOk.Caption = constValider
                            cmdSave.Caption = constàModifier
         
            Case Is = "? "
                        If Fct = "Delete" Then
                            cmdSave.Caption = constDelete
                        Else
                            cmdOk.Caption = constàValider
                        End If
        End Select
    Else
        fgEchéancier_Load
    End If
    cmdSave.Visible = blncmdSave_Visible
    blnfgSelect_DisplayLine = True
End If

End Sub

Public Sub fgEchéancier_AddNew()
V = srvTFlux_Dtaq_Put("Init", recTFlux)
If Not IsNull(V) Then fgEchéancier_Delete: Exit Sub
For I = 1 To arrTFlux_Nb
    arrTFlux(I).Method = constAddNew
    arrTFlux(I).IdRéférence = recTope.IdRéférence
    recTFlux = arrTFlux(I)
    If recTope.optReprise = "R" And recTFlux.AmjEchéanceTrt < DSys Then
        recTFlux.Statut = "R"
        recTFlux.StatutPlus = "ep"
    End If
     arrTFlux(I) = recTFlux
    V = srvTFlux_Dtaq_Put("Add", recTFlux)
    If Not IsNull(V) Then fgEchéancier_Delete: Exit Sub
Next I
V = srvTFlux_Dtaq_Put("Snd", recTFlux)
If Not IsNull(V) Then fgEchéancier_Delete: Exit Sub

End Sub

Public Sub fgEchéancier_Update()
V = srvTFlux_Dtaq_Put("Init", arrTFlux(1))
If Not IsNull(V) Then: Exit Sub
For I = 1 To arrTFlux_Nb
    If arrTFlux(I).Method = constUpdate Or arrTFlux(I).Method = constAddNew Then
        V = srvTFlux_Dtaq_Put("Add", arrTFlux(I))
        If Not IsNull(V) Then Exit Sub
    End If
Next I
V = srvTFlux_Dtaq_Put("Snd", arrTFlux(1))
If Not IsNull(V) Then Exit Sub

End Sub

Public Sub fgEchéancier_Delete()
Call lstErr_AddItem(lstErr, cmdContext, V)
Call lstErr_AddItem(lstErr, cmdContext, "!! Suppression de l'échéancier")
arrTFlux(1).Method = "DeleteAll"
Call srvTFlux_Update(arrTFlux(1))
End Sub

Public Sub fgEchéancier_Load()
ReDim arrTFlux(1)

recTFlux_Init recTFlux
recTFlux.Method = "SnapP0"
recTFlux.IdRéférence = mTope.IdRéférence

arrTFlux(0) = recTFlux
arrTFlux(0).IdSéquence = 999

Call srvTFlux_Load(recTFlux, arrTFlux(0))
arrTFlux_Nb = srvTFlux.arrTFlux_Nb
ReDim arrTFlux(arrTFlux_Nb)
For I = 1 To arrTFlux_Nb
    arrTFlux(I) = srvTFlux.arrTFlux(I)
Next I

fgEchéancier_Display "T"

End Sub

Public Sub cmdSave_àCompta()

frmPrêt.Enabled = False

blnErr = False
mTflux = arrTFlux(1)
mTflux.Method = constàCompta
V = srvTFlux_Update(mTflux)
If Not IsNull(V) Then MsgBox "frmPrêt cmdSaveàCompta : recherche numlot": GoTo Exit_Sub
mTflux.CptMvtAMJ = DSys
mTflux.CptMvtHMS = time_Hms

V = srvTFlux_Dtaq_Put("Init", arrTFlux(1))

If IsNull(V) Then
    For I = 1 To arrTFlux_Nb
        arrTFlux(I).Method = "Update"
        arrTFlux(I).Statut = "à"
        arrTFlux(I).StatutPlus = "C "
        arrTFlux(I).CptMvtLot = mTflux.CptMvtLot
        arrTFlux(I).CptMvtPièce = I
        arrTFlux(I).CptMvtAMJ = mTflux.CptMvtAMJ
        arrTFlux(I).CptMvtHMS = mTflux.CptMvtHMS
        V = srvTFlux_Dtaq_Put("Add", arrTFlux(I))
        If Not IsNull(V) Then blnErr = True: GoTo Exit_Sub
    Next I
    V = srvTFlux_Dtaq_Put("Snd", arrTFlux(1))
    If Not IsNull(V) Then blnErr = True: GoTo Exit_Sub
End If

If Not blnErr Then
    lastActiveControl_Name = ""
    cmdOk.Visible = False
    cmdSave.Visible = False
    Call lstErr_AddItem(lstErr, cmdContext, "àCompta - N° lot :  : " & mTflux.CptMvtLot)
    If Not blnComptaAuto Then TFlux_Compta.LotàCompta_Demande mTflux.CptMvtLot
Else
    Call lstErr_AddItem(lstErr, cmdContext, V)
    cmdReset
    MsgBox "frmPrêt cmdSaveàCompta : à faire annuler demande de comptabilisation"
End If
fgEchéancier.Clear: fgEchéancier.Rows = 1


Exit_Sub:

frmPrêt.Enabled = True
AppActivate frmPrêt.Caption

End Sub

Public Sub fgSelect_DisplayLine()
fgSelect_K = (fgSelect.Row) * fgSelect.Cols
fgSelect.TextArray(0 + fgSelect_K) = ""
fgSelect.TextArray(0 + fgSelect_K) = recStatut_Libellé(arrTOpe(arrTOpe_Index).Statut & arrTOpe(arrTOpe_Index).StatutPlus)
fgSelect.TextArray(1 + fgSelect_K) = arrTOpe(arrTOpe_Index).Nature
fgSelect.TextArray(2 + fgSelect_K) = Compte_Imp(arrTOpe(arrTOpe_Index).EngagementCompte)
Call CV_AttributS(arrTOpe(arrTOpe_Index).Devise, CV1)
recCompteInit recCompte
recCompte.Société = SocId$
recCompte.Agence = SocAgence$
recCompte.Devise = CV1.DeviseN
recCompte.Numéro = arrTOpe(arrTOpe_Index).EngagementCompte
mdbCptP0_Find recCompte
fgSelect.TextArray(3 + fgSelect_K) = recCompte.Intitulé
fgSelect.TextArray(4 + fgSelect_K) = Format(arrTOpe(arrTOpe_Index).Capital, "#### ### ###.00 ") & arrTOpe(arrTOpe_Index).Devise
fgSelect.TextArray(5 + fgSelect_K) = Format(arrTOpe(arrTOpe_Index).TauxMarge, "#0.00000")
fgSelect.TextArray(6 + fgSelect_K) = arrTOpe(arrTOpe_Index).RéférenceInterne
fgSelect.TextArray(7 + fgSelect_K) = arrTOpe(arrTOpe_Index).RéférenceExterne
fgSelect.TextArray(8 + fgSelect_K) = arrTOpe(arrTOpe_Index).MajUsr & " " & dateImp(arrTOpe(arrTOpe_Index).MajAMJ) & " " & timeImp(arrTOpe(arrTOpe_Index).MajHMS)
fgSelect.TextArray(9 + fgSelect_K) = arrTOpe(arrTOpe_Index).ValUsr & " " & dateImp(arrTOpe(arrTOpe_Index).valAMJ) & " " & timeImp(arrTOpe(arrTOpe_Index).ValHMS)
fgSelect.TextArray(10 + fgSelect_K) = Format(arrTOpe(arrTOpe_Index).IdRéférence, "#### ### ##0 ")
fgSelect.TextArray(11 + fgSelect_K) = arrTOpe_Index

End Sub

Public Sub txtAmjEchéance_Init()
If Not blnAmjEchéance Then
    wAmjEchéance = dateFinDeMois(dateElp(fctPériodicité, 1, wAmjEngagement))
    Call DTPicker_Set(txtAMJEchéance, wAmjEchéance)
End If
End Sub

Public Sub cmdControl_Compte(xMsg As String)
Dim X As String
If xMsg < 100000 Then
    xMsg = xMsg * 1000000 + 10
    Call Compte_BiaTyp(xMsg, paramTFlux_BiatypEchéance)
End If

X = xMsg
If Trim(txtEngCompte) = "" Then
    Call Compte_BiaTyp(X, paramTFlux_BiatypEngagement)
    txtEngCompte = X
End If

If Trim(txtCorrCompte) = "" Then
    If paramTFlux_CompteDéblocageDesFonds = "00000000000" Then
        Call Compte_BiaTyp(X, paramTFlux_BiatypEngagementCorr)
        txtCorrCompte = X
    Else
        txtCorrCompte = paramTFlux_CompteDéblocageDesFonds
    End If
    
End If

End Sub

Public Sub mnuComptaEchéancier_Load()
ReDim arrTFlux(1)

recTFlux_Init recTFlux
recTFlux.Method = "SnapLE"
recTFlux.AmjEchéanceTrt = "00000000"

arrTFlux(0) = recTFlux
arrTFlux(0).AmjEchéanceTrt = wAmjEchéanceTrt
arrTFlux(0).IdRéférence = 999999999
arrTFlux(0).IdSéquence = 32000

Call srvTFlux_Load(recTFlux, arrTFlux(0))
arrTFlux_Nb = srvTFlux.arrTFlux_Nb
ReDim arrTFlux(arrTFlux_Nb)
For I = 1 To arrTFlux_Nb
    arrTFlux(I) = srvTFlux.arrTFlux(I)
Next I
SSTab1.Tab = 3
fgEchéancier_Display " "

If arrTFlux_Nb = 0 Then
    MsgBox "frmPrêt mnuComptaEchéancier : PAS D'ECHEANCE A TRAITER"
Else
    cmdOk.Caption = constàCompta: cmdOk.Visible = True
    currentAction = constàCompta
End If

End Sub

Public Sub fgEchéancier_RemboursementAnticipé()
Dim curX As Currency, I As Integer, I1 As Integer, wNbj As Long

curX = 0: I1 = 0
recTFlux = arrTFlux(arrTFlux_Nb)

For I = 1 To arrTFlux_Nb
   
    If arrTFlux(I).CodeOpération = "PR02" And arrTFlux(I).Statut = " " Then
        If I1 = 0 Then I1 = I
        curX = curX + arrTFlux(I).Capital
        arrTFlux(I).Method = constUpdate
        arrTFlux(I).Statut = "A"
        arrTFlux(I).StatutPlus = "RA"
    End If
Next I

With recTFlux                                   ' RemboursementAnticipé
    .Method = constAddNew
    .IdSéquence = recTFlux.IdSéquence + 1
    .CodeOpération = "PR05"
    .Capital = curX
    .Intérêts = 0
    .Taux = recTope.TauxMarge
    .AmjEchéanceTrt = DSys
    .AmjDébut = wAmjRemboursementAnticipé
    .AmjFin = wAmjRemboursementAnticipé
    .AmjOpération = DSys
    .AmjValeur = wAmjRemboursementAnticipé
End With
ReDim Preserve arrTFlux(arrTFlux_Nb + 2)
arrTFlux_Nb = arrTFlux_Nb + 1
arrTFlux(arrTFlux_Nb) = recTFlux

If I1 > 0 Then
    xTOpe = recTope
    With xTOpe
        .Capital = curX
        .AmjDébut = arrTFlux(I1).AmjDébut
        .AmjFin = wAmjRemboursementAnticipé
    End With
    V = fctTOpe_Intérêts(xTOpe, CV1, curX, wNbj)
    
    If IsNull(V) Then
        If curX <> 0 Then
            With recTFlux                                   ' intérêts décalés
                .IdSéquence = recTFlux.IdSéquence + 1
                .CodeOpération = "PR04"
                .Capital = 0
                .Intérêts = curX
                .Nbj = wNbj
                .AmjDébut = arrTFlux(I1).AmjDébut
            End With
            arrTFlux_Nb = arrTFlux_Nb + 1
            arrTFlux(arrTFlux_Nb) = recTFlux
        End If
    End If
End If

End Sub

Public Sub cmdRemboursementAnticipé_àCompta()
ReDim arrTFlux(1)

recTFlux_Init recTFlux
recTFlux.Method = "SnapP0"
recTFlux.IdRéférence = recTope.IdRéférence

arrTFlux(0) = recTFlux
arrTFlux(0).IdSéquence = 32000

Call srvTFlux_Load(recTFlux, arrTFlux(0))
ReDim arrTFlux(srvTFlux.arrTFlux_Nb)
arrTFlux_Nb = 0

For I = 1 To srvTFlux.arrTFlux_Nb
    If srvTFlux.arrTFlux(I).Statut = " " Then
        arrTFlux_Nb = arrTFlux_Nb + 1
        arrTFlux(arrTFlux_Nb) = srvTFlux.arrTFlux(I)
    End If
Next I
SSTab1.Tab = 3
fgEchéancier_Display " "

If arrTFlux_Nb > 0 Then cmdSave_àCompta

End Sub
