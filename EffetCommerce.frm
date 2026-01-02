VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEffetCommerce 
   AutoRedraw      =   -1  'True
   Caption         =   "Effet de commerce"
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
      TabIndex        =   34
      Top             =   0
      Width           =   3500
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   0
      TabIndex        =   27
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Menu"
      TabPicture(0)   =   "EffetCommerce.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraOption"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Sélection"
      TabPicture(1)   =   "EffetCommerce.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgSelect"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Opération 1/ 2"
      TabPicture(2)   =   "EffetCommerce.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraOpération1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Opération 2/2"
      TabPicture(3)   =   "EffetCommerce.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraOpération2"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Flux"
      TabPicture(4)   =   "EffetCommerce.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picGFlux"
      Tab(4).Control(1)=   "fgFlux"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Echéancier"
      TabPicture(5)   =   "EffetCommerce.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "picGEch"
      Tab(5).Control(1)=   "fgEch"
      Tab(5).ControlCount=   2
      Begin VB.Frame fraOpération2 
         Caption         =   "Caractéristiques de l'effet 2/2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5655
         Left            =   -74880
         TabIndex        =   50
         Top             =   360
         Width           =   9135
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
            Height          =   615
            Left            =   7800
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   4200
            Width           =   1200
         End
         Begin VB.CommandButton cmdOpérationElpDisplay 
            Caption         =   "afficher"
            Height          =   375
            Left            =   7800
            TabIndex        =   24
            Top             =   3720
            Width           =   1215
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
            Height          =   615
            Left            =   7800
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   4920
            Width           =   1200
         End
         Begin VB.Frame fraConditions 
            Caption         =   "Conditions"
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
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   8895
            Begin VB.Frame fraConditions_Escompte 
               Caption         =   "Escompte"
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
               Left            =   120
               TabIndex        =   61
               Top             =   1320
               Width           =   8655
               Begin VB.TextBox txtTauxNonAccepté 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3360
                  TabIndex        =   21
                  Top             =   1320
                  Width           =   1215
               End
               Begin VB.CheckBox chkTauxNonAccepté 
                  Caption         =   "% majoration effet non accepté"
                  CausesValidation=   0   'False
                  Height          =   375
                  Left            =   120
                  TabIndex        =   20
                  Top             =   1320
                  Width           =   2535
               End
               Begin VB.TextBox txtTauxMajoré 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3360
                  TabIndex        =   19
                  Top             =   840
                  Width           =   1215
               End
               Begin VB.CheckBox chkTauxMajoré 
                  Caption         =   "% majoration du taux (>  90 jours)"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   18
                  Top             =   840
                  Width           =   2775
               End
               Begin VB.TextBox txtTaux 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3360
                  TabIndex        =   17
                  Top             =   360
                  Width           =   1215
               End
               Begin VB.Label lblTaux 
                  Caption         =   "taux %"
                  Height          =   255
                  Left            =   480
                  TabIndex        =   63
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label libIntérêts 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "-"
                  Height          =   315
                  Left            =   5040
                  TabIndex        =   62
                  Top             =   360
                  Width           =   1395
               End
            End
            Begin VB.CheckBox chkComEndos 
               Caption         =   "commission df'endos ( pour mille)"
               Height          =   255
               Left            =   240
               TabIndex        =   15
               Top             =   840
               Width           =   2655
            End
            Begin VB.CheckBox chkComManipulation 
               Caption         =   "commission de manipulation (par effet)"
               Height          =   255
               Left            =   240
               TabIndex        =   13
               Top             =   360
               Width           =   3255
            End
            Begin VB.TextBox txtComManipulation 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3480
               TabIndex        =   14
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox txtComEndos 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3480
               TabIndex        =   16
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label libTVA 
               Caption         =   "TVA"
               Height          =   255
               Left            =   7080
               TabIndex        =   60
               Top             =   360
               Width           =   495
            End
            Begin VB.Label libFrais3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               Height          =   315
               Left            =   7560
               TabIndex        =   59
               Top             =   240
               Width           =   1155
            End
            Begin VB.Label libFrais2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               Height          =   315
               Left            =   5160
               TabIndex        =   58
               Top             =   240
               Width           =   1395
            End
            Begin VB.Label libFrais1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               Height          =   315
               Left            =   5160
               TabIndex        =   57
               Top             =   720
               Width           =   1395
            End
         End
         Begin VB.Label libStatut 
            BorderStyle     =   1  'Fixed Single
            Height          =   1335
            Left            =   120
            TabIndex        =   53
            Top             =   4200
            Width           =   3495
         End
         Begin VB.Label libInfo 
            BorderStyle     =   1  'Fixed Single
            Height          =   1335
            Left            =   3720
            TabIndex        =   52
            Top             =   4200
            Width           =   3975
         End
      End
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
         TabIndex        =   48
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
         TabIndex        =   47
         Top             =   5400
         Width           =   9045
      End
      Begin VB.Frame fraOpération1 
         Caption         =   "Caractéristiques de l'effet 1/2"
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
         Left            =   -74880
         TabIndex        =   35
         Top             =   360
         Width           =   9135
         Begin VB.Frame fraTiré 
            Caption         =   "Tiré"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   120
            TabIndex        =   64
            Top             =   2760
            Width           =   8895
            Begin VB.TextBox txtTiréRéférence 
               Height          =   285
               Left            =   1800
               TabIndex        =   9
               Top             =   1200
               Width           =   6615
            End
            Begin VB.TextBox txtTiréDomiciliation 
               Height          =   285
               Left            =   1800
               TabIndex        =   8
               Top             =   720
               Width           =   6615
            End
            Begin VB.TextBox txtTiréNom 
               Height          =   285
               Left            =   1800
               TabIndex        =   7
               Top             =   240
               Width           =   6615
            End
            Begin MSComCtl2.DTPicker txtAmjMCNE 
               Height          =   300
               Left            =   1800
               TabIndex        =   10
               Top             =   1650
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
               Format          =   64880643
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblTiréAMJExpédition 
               Caption         =   "MCNE date expédition"
               Height          =   255
               Left            =   120
               TabIndex        =   68
               Top             =   1680
               Width           =   1695
            End
            Begin VB.Label lblTiréRéférence 
               Caption         =   "Référence facture ..."
               Height          =   255
               Left            =   120
               TabIndex        =   67
               Top             =   1200
               Width           =   1455
            End
            Begin VB.Label lblTiréDomiciliation 
               Caption         =   "domiciliation du tiré"
               Height          =   255
               Left            =   120
               TabIndex        =   66
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label lblTiréNom 
               Caption         =   "nom du tiré"
               Height          =   255
               Left            =   120
               TabIndex        =   65
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.TextBox txtEngagementCompte 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1920
            TabIndex        =   2
            Top             =   840
            Width           =   1095
         End
         Begin VB.ComboBox cboNature 
            ForeColor       =   &H00FF0000&
            Height          =   315
            ItemData        =   "EffetCommerce.frx":00A8
            Left            =   1920
            List            =   "EffetCommerce.frx":00AA
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   360
            Width           =   4575
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
            Left            =   1920
            TabIndex        =   3
            Top             =   1320
            Width           =   2055
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
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1320
            Width           =   1400
         End
         Begin VB.TextBox txtRéférenceExterne 
            Height          =   285
            Left            =   6120
            MaxLength       =   16
            TabIndex        =   12
            Top             =   5040
            Width           =   2655
         End
         Begin VB.TextBox txtRéférenceInterne 
            Height          =   285
            Left            =   1680
            MaxLength       =   16
            TabIndex        =   11
            Top             =   5040
            Width           =   2655
         End
         Begin VB.CheckBox chkComptaReprise 
            Alignment       =   1  'Right Justify
            Caption         =   "Reprise de l'en-cours"
            Height          =   255
            Left            =   6960
            TabIndex        =   36
            Top             =   240
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker txtAmjEngagement 
            Height          =   300
            Left            =   1920
            TabIndex        =   5
            Top             =   1800
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
            Format          =   64880643
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin MSComCtl2.DTPicker txtAMJEchéance 
            Height          =   300
            Left            =   1920
            TabIndex        =   6
            Top             =   2280
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
            Format          =   64880643
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label libEngagementCompte 
            BackColor       =   &H00E0E0E0&
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
            Left            =   3360
            TabIndex        =   56
            Top             =   840
            Width           =   5655
         End
         Begin VB.Label libCompte2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            Height          =   315
            Left            =   4320
            TabIndex        =   55
            Top             =   2280
            Width           =   1395
         End
         Begin VB.Label libCompte1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            Height          =   315
            Left            =   4320
            TabIndex        =   54
            Top             =   1800
            Width           =   1395
         End
         Begin VB.Label lblAMJEchéance 
            Caption         =   "Echéance"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label lblRéférenceExterne 
            Caption         =   "Réf bordereau"
            Height          =   255
            Left            =   4800
            TabIndex        =   44
            Top             =   5040
            Width           =   1215
         End
         Begin VB.Label lblRéférenceInterne 
            Caption         =   "Référence service"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   5040
            Width           =   1575
         End
         Begin VB.Label lblAmjDébut 
            Caption         =   "Date de remise"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label lblCapital 
            Caption         =   "Montant"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label lblEngagementCompte 
            Caption         =   "compte d'effets"
            Height          =   255
            Left            =   6240
            TabIndex        =   40
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lblBénéficiaire 
            Caption         =   "Bénéficiaire"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblNature 
            Caption         =   "Nature"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblEchCompte 
            Caption         =   "Compte à créditer"
            Height          =   255
            Left            =   6240
            TabIndex        =   37
            Top             =   2280
            Width           =   1335
         End
      End
      Begin VB.Frame fraOption 
         Caption         =   "Options"
         Height          =   4455
         Left            =   600
         TabIndex        =   29
         Top             =   840
         Width           =   7935
         Begin VB.CommandButton cmdSelect 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Rechercher"
            Height          =   975
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   2760
            Width           =   2415
         End
         Begin VB.OptionButton optSelectX 
            Caption         =   "à faire"
            Height          =   255
            Left            =   360
            TabIndex        =   69
            Top             =   2040
            Value           =   -1  'True
            Width           =   3135
         End
         Begin VB.TextBox txtSelect 
            Height          =   285
            Left            =   5760
            TabIndex        =   30
            Top             =   600
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker txtAmjMax 
            Height          =   300
            Left            =   6480
            TabIndex        =   31
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
            Format          =   64880643
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblSelect 
            Caption         =   "Client"
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
            Left            =   4440
            TabIndex        =   33
            Top             =   600
            Width           =   975
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
            Left            =   3960
            TabIndex        =   32
            Top             =   1440
            Width           =   2295
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSelect 
         Height          =   5250
         Left            =   -74880
         TabIndex        =   28
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
         FormatString    =   $"EffetCommerce.frx":00AC
      End
      Begin MSFlexGridLib.MSFlexGrid fgEch 
         Height          =   5010
         Left            =   -74880
         TabIndex        =   45
         Top             =   360
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
         FormatString    =   $"EffetCommerce.frx":01D4
      End
      Begin MSFlexGridLib.MSFlexGrid fgFlux 
         Height          =   4770
         Left            =   -74880
         TabIndex        =   46
         Top             =   480
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
         FormatString    =   $"EffetCommerce.frx":02C7
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8880
      Picture         =   "EffetCommerce.frx":03D2
      Style           =   1  'Graphical
      TabIndex        =   25
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
      TabIndex        =   26
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
      Begin VB.Menu Validation 
         Caption         =   "-"
      End
      Begin VB.Menu mnuValidationDemande 
         Caption         =   "Demande de validation (bordereau)"
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
         Caption         =   "Lot à comptabiliser : invalider"
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
Attribute VB_Name = "frmEffetCommerce"
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
Dim EffetCommerceAut As typeAuthorization
Dim blnCompteSituation_Saisie As Boolean, blnCompteSituation_Validation As Boolean, blnCompteSituation_Forçage As Boolean

Dim meElpTable As typeElpTable, meCompte As typeCompte, meRacine As typeRacine

Dim wAmjEngagement As String, wAmjEchéance As String, blnAmjEchéance As Boolean
Dim wAmjDébut  As String, wAmjFin As String
Dim paramAmjEngagementMin As String, paramAmjEngagementMax As String
Dim paramAmjEchéanceMin As String, paramAmjEchéanceMax As String
Dim wAMJEffet  As String, wAMJValeur As String
Dim meCV1 As typeCV, meCV2 As typeCV, meCV3 As typeCV

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

Dim meGOpe As typeGOpe, xGOpe As typeGOpe, mGOpe As typeGOpe, mEchéancierGOpe As typeGOpe

Dim mearrGOpe() As typeGOpe
Dim mearrGOpe_Nb As Integer, mearrGOpe_Index As Integer, mearrGOpe_NbMax As Integer

Dim mearrGFlux() As typeGFlux, meGFlux As typeGFlux, mGFlux As typeGFlux
Dim mearrGFlux_Nb As Integer, mearrGFlux_Index As Integer, mearrGFlux_NbMax As Integer
'''Dim saveGFlux() As typeGFlux, saveGFlux_Index As Integer, saveGFlux_Nb As Integer

Dim mearrGEch() As typeGEch, meGECh As typeGEch, mGECh As typeGEch
Dim mearrGEch_Nb As Integer, mearrGEch_Index As Integer, mearrGEch_NbMax As Integer
Dim warrGEch() As typeGEch, warrGEch_Nb As Integer

Dim mearrGMemo() As typegMemo, meGMemo As typegMemo, mGMemo As typegMemo
Dim mearrGMemo_Nb As Integer, mearrGMemo_Index As Integer, mearrGMemo_NbMax As Integer
Dim warrGMemo() As typegMemo, warrGMemo_Nb As Integer, warrGMemo_NbMax As Integer
Dim xarrGMemo() As typegMemo, xarrGMemo_Nb As Integer, xarrGMemo_NbMax As Integer

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


Dim paramEffetCommerce As typeGParam
Dim mtxtEngagementCompte As String, mcboNature_ListIndex As Integer
Dim wTaux As Double

Dim blnTauxMajoré_Set As Boolean, blnTauxMajoré_MsgBox As Boolean

Dim mLotàComptabiliserValider As String, mLotàComptabiliserAnnuler As String, mLotàComptabiliserPrint As String
Public Function Compte_Load(lDevise As String, lCompteNuméro As String)
Dim V

V = Null

If lCompteNuméro <> meCompte.Numéro Then
    meCompte.Société = SocId$
    meCompte.Agence = SocAgence$
    meCompte.Devise = lDevise
    meCompte.Numéro = lCompteNuméro
    If blnJPL Then
        V = mdbCptP0_Find(meCompte)
    Else
        V = srvCompte_InitFind(meCompte)
    End If
    If Not IsNull(V) Then
        blnCompteSituation_Saisie = False
        recCompteInit meCompte
        V = "? compte inconnu : " & meCompte.Devise & " _" & meCompte.Numéro
    Else
        Select Case currentAction
            Case constValider: V = CompteSituation_Validation(meCompte, blnCompteSituation_Validation, blnCompteSituation_Forçage)
            Case constSaisie: V = CompteSituation_Saisie(meCompte, blnCompteSituation_Saisie)
        End Select
    End If
    If Not IsNull(V) Then meCompte.Numéro = "$$$" ''': Call lstErr_AddItem(lstErr, lstErr, "? " & V):

End If
Compte_Load = V

End Function


Public Sub fraOpération_Load(Fct As String)
Dim X As String

fgSelect_RowClick = 0
Call fgSelect_Color(fgSelect_RowDisplay, vbCyan, fgSelect_ColorClick) 'txtUsr.BackColor)
blnControl = False
mGOpe.Method = "SeekP0"
V = srvGOpe_Monitor(mGOpe)
If IsNull(V) Then
    GEch_Load
    GFlux_Load
    GMemo_Load
    meGOpe = mGOpe
        
'22011106 V = srvGSub.GOpération_Load(paramEffetCommerce, meGOpe, mearrGEch_Nb, mearrGEch(), mearrGFlux_Nb, mearrGFlux(), mearrGMemo_Nb, mearrGMemo())
'22011106 If IsNull(V) Then

    lstErr.Clear: lstErr.Height = 200
    blnAmjEchéance = True
    SSTab1.Tab = 2
    mGOpe = xGOpe
    mGOpe.Method = Fct
        
    If mearrGEch_Nb = 1 Then
        ReDim Preserve mearrGEch(10)
        Call GEch_Gen(paramEffetCommerce, meGOpe, mearrGEch_Nb, mearrGEch())
    End If
    fgEch_Display
    
    If mearrGFlux_Nb = 0 Then Call GFlux_Gen(paramEffetCommerce, meGOpe, mearrGFlux_Nb, mearrGFlux())
    fgFlux_Display

    
    cbo_Scan mGOpe.Nature, cboNature
    
    paramEffetCommerce.NatureCode = mGOpe.Nature
    V = srvGSub.param_Nature(paramEffetCommerce)
    txtEngagementCompte = mId$(mGOpe.EngagementCompte, 1, 5)
        
    cmdControl_OpérationG
    
    meCV1.DeviseIso = mGOpe.Devise1
    V = CV_Attribut(meCV1): If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, "? erreur devise : " & meCV1.DeviseIso)
    cbo_Scan mGOpe.Devise1, cboDevise1

    libCompte1 = Compte_Imp(mGOpe.EngagementCompte)
    libCompte2 = Compte_Imp(mGOpe.EchéanceCompte)
    
    txtMontant1 = Format$(mGOpe.Montant1, "### ### ### ##0.00")
      
    Call DTPicker_Set(txtAmjEngagement, mGOpe.AmjDébut): wAmjEngagement = mGOpe.AmjEngagement
    Call DTPicker_Set(txtAMJEchéance, mGOpe.AmjFin)
    
    txtTiréNom = mId$(mearrGMemo(1).MemoText, 1, 50)

    txtTiréDomiciliation = mId$(mearrGMemo(1).MemoText, 51, 50)

    txtTiréRéférence = mId$(mearrGMemo(1).MemoText, 101, 50)

    Call DTPicker_Set(txtAmjMCNE, mId$(mearrGMemo(1).MemoText, 151, 8))

    txtRéférenceInterne = mGOpe.RéférenceInterne
    'txtPréavisNbj = meGOpe.PréavisNbj
    
    txtTaux = Format$(mGOpe.TauxMarge1, "#####0.00000")

    currentAction = constDisplay
    fraOpération_Display

    If mGOpe.Statut = "à" Then
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

recGEch_Init meGECh

meGECh.Method = "SnapP0"
meGECh.IdRéférence = mGOpe.IdRéférence
mearrGEch(0) = meGECh
mearrGEch(0).EchSéquence = 99999
Call srvGEch_Load(meGECh, mearrGEch(0))

mearrGEch_Nb = srvGEch.arrGECh_Nb
mearrGEch_NbMax = mearrGEch_Nb + 1: ReDim mearrGEch(mearrGEch_NbMax)

For I = 1 To mearrGEch_Nb
    mearrGEch(I) = srvGEch.arrGECh(I)
    mearrGEch(I).Method = ""
Next I

End Sub

Public Function GEch_Save(lIdRéférence As Long)
Dim I As Integer

GEch_Save = Null

For I = 1 To mearrGEch_Nb
    If mearrGEch(I).Method = constAddNew Or mearrGEch(I).Method = constUpdate Then
        If mearrGEch(I).IdRéférence = 0 Then mearrGEch(I).IdRéférence = lIdRéférence
        If mearrGEch(I).IdRéférence = lIdRéférence Then
            V = srvGEch_Update(mearrGEch(I))
            If Not IsNull(V) Then GEch_Save = V
        End If
    End If
Next I

End Function


Public Function GFlux_Save(lIdRéférence As Long)
Dim I As Integer

GFlux_Save = Null

For I = 1 To mearrGFlux_Nb
    If mearrGFlux(I).Method = constAddNew Or mearrGFlux(I).Method = constUpdate Then
        If mearrGFlux(I).IdRéférence = 0 Then mearrGFlux(I).IdRéférence = lIdRéférence
        If mearrGFlux(I).IdRéférence = lIdRéférence Then
            V = srvGFlux_Update(mearrGFlux(I))
            If Not IsNull(V) Then GFlux_Save = V
        End If
    End If
Next I

End Function

Public Function GMemo_Save(lIdRéférence As Long)
Dim I As Integer

GMemo_Save = Null

For I = 1 To mearrGMemo_Nb
    If mearrGMemo(I).Method = constAddNew Or mearrGMemo(I).Method = constUpdate Then
        If mearrGMemo(I).IdRéférence = 0 Then mearrGMemo(I).IdRéférence = lIdRéférence
        If mearrGMemo(I).IdRéférence = lIdRéférence Then
            V = srvGMemo_Update(mearrGMemo(I))
            If Not IsNull(V) Then GMemo_Save = V
        End If
    End If
Next I

End Function


Public Sub GFlux_Load()
Dim I As Integer

recGFlux_Init meGFlux

meGFlux.Method = "SnapP0"
meGFlux.IdRéférence = mGOpe.IdRéférence
mearrGFlux(0) = meGFlux
mearrGFlux(0).FluxSéquence = 99999
Call srvGFlux_Load(meGFlux, mearrGFlux(0))

mearrGFlux_Nb = srvGFlux.arrGFlux_Nb
mearrGFlux_NbMax = mearrGFlux_Nb + 1: ReDim mearrGFlux(mearrGFlux_NbMax)

For I = 1 To mearrGFlux_Nb
    mearrGFlux(I) = srvGFlux.arrGFlux(I)
    mearrGFlux(I).Method = ""
Next I

End Sub

Public Sub GMemo_Load()
Dim I As Integer

recGMemo_Init meGMemo

meGMemo.Method = "SnapP0"
meGMemo.IdRéférence = mGOpe.IdRéférence
mearrGMemo(0) = meGMemo
mearrGMemo(0).MemoSéquence = 99999
Call srvGMemo_Load(meGMemo, mearrGMemo(0))

mearrGMemo_Nb = srvGMemo.arrgMemo_NB
mearrGMemo_NbMax = mearrGMemo_Nb + 1: ReDim mearrGMemo(mearrGMemo_NbMax)

For I = 1 To mearrGMemo_Nb
    mearrGMemo(I) = srvGMemo.arrgMemo(I)
    mearrGMemo(I).Method = ""
Next I

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
    cmdOk.Visible = False
    cmdSave.Visible = False
    currentAction = ""
    cmdContext.Caption = constcmdRechercher
    fgSelect.Enabled = True
    fgFlux.Enabled = True
    fraOpération1.Enabled = False
    fraOpération2.Enabled = False
    If fgSelect.Rows > 1 Then
        SSTab1.Tab = 1
    Else
        cmdReset
    End If
End If

End Sub
Public Sub cmdControl()

If Not Me.Enabled Then Exit Sub
If SSTab1.Tab <> 2 And SSTab1.Tab <> 3 Then Exit Sub
Me.Enabled = False

cmdOk.Visible = False
cmdSave.Visible = False
blnControl = False
blnSetfocus = False

lstErr.Clear
lstErr.Height = 200
libRéférenceInterne = currentAction & " : " & mGOpe.Nature & " : " & mGOpe.RéférenceInterne
lastActiveControl_Name = currentActiveControl_Name
xGOpe = meGOpe
meGOpe = mGOpe



Select Case currentAction
    Case constSaisie:
            meGOpe.Application = paramEffetCommerce.Application
            meGOpe.IPA = "A"
            meGOpe.NbjBase = "0"

            mearrGEch(1).IdRéférence = 0
            blnControlBiatyp = True: cmdControl_OpérationG
            If lstErr.ListCount = 0 Then
                cmdControl_OpérationD
                Call GFlux_Gen(paramEffetCommerce, meGOpe, mearrGFlux_Nb, mearrGFlux())
                fgFlux_Display
                Call GEch_Gen(paramEffetCommerce, meGOpe, mearrGEch_Nb, mearrGEch())
                fgEch_Display
            End If

    Case constDisplay
            cmdControl_OpérationD
            currentActiveControl_Name = ""
     Case constValider
            cmdControl_OpérationD
            currentActiveControl_Name = ""
           V = fctGOpe_Compare(meGOpe, mGOpe)
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
    
If blnSetfocus And lastActiveControl_Name <> currentActiveControl_Name Then
    Select Case currentActiveControl_Name
        Case "cmdOk": cmdOk.SetFocus
        Case "txtRéférenceInterne": txtRéférenceInterne.SetFocus
        Case "txtEngagementCompte": txtEngagementCompte.SetFocus
        Case "txtMontant1": txtMontant1.SetFocus
        Case "txtTaux": txtTaux.SetFocus
        Case "txtTiréDomiciliation": txtTiréDomiciliation.SetFocus
        Case "txtTiréNom": txtTiréNom.SetFocus
        Case "txtTiréRéférence": txtTiréRéférence.SetFocus
   End Select
End If

blnControl = True

End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
picGEch.BackColor = greenColor.BackColor
picGFlux.BackColor = greenColor.BackColor

recRacineInit meRacine
cmdOk.Caption = constàValider: cmdOk.Visible = False
cmdSave.Caption = constEnAttente: cmdSave.Visible = False
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
blnComptaAuto = False
cmdOk.FontSize = 8: cmdOk.FontName = "MS Sans Serif"
blncmdOk_Visible = False: blncmdSave_Visible = False
blnfgSelect_DisplayLine = False: blnfgEchéance_DisplayLine = False
    
cmdOpérationElpDisplay.Visible = EffetCommerceAut.Xspécial
'If EffetCommerceAut.Xspécial Then
'    cmdSave.Height = cmdOk.Height / 2
'Else
'    cmdSave.Height = cmdOk.Height
'End If

fgFlux.Clear: fgFlux.Rows = 1: fgFlux_RowDisplay = 0
fgEch.Clear: fgEch.Rows = 1: fgEch_RowDisplay = 0
If cboNature.ListCount > 0 Then cboNature.ListIndex = 0
If cboDevise1.ListCount > 0 Then cboDevise1.ListIndex = 0
mtxtEngagementCompte = Chr$(9): mcboNature_ListIndex = -2

txtMontant1 = ""
txtTaux = ""
mAMJReprise = DSys
wAmjEchéance = dateFinDeMois(dateElp("MoisAdd", 1, DSys)): Call DTPicker_Set(txtAMJEchéance, wAmjEchéance)
txtRéférenceInterne = "": txtRéférenceExterne = ""
txtEngagementCompte = "": libEngagementCompte = ""
libStatut = ""
libFrais1 = "": libFrais2 = "": libFrais3 = ""
txtTiréNom = "": txtTiréDomiciliation = "": txtTiréRéférence = ""

mGOpe.Statut = "à"
mGOpe.StatutPlus = "?"
mGOpe.Method = constAddNew
mEChéanceCompte = Space$(11): mEngagementCompte = Space$(11): mEngagementCorrCompte = Space$(11)
Call DTPicker_Set(txtAmjEngagement, DSys)
Call DTPicker_Set(txtAMJEchéance, DsysMinus2) 'DValNext2)
Call DTPicker_Set(txtAmjMCNE, DSys)
'chkComptaReprise = "0"

fraOption.Visible = True
blnEchéancier_Gen = False
fraOpération1.Enabled = False: fraOpération2.Enabled = False
lastActiveControl_Name = "": currentActiveControl_Name = ""

picGEch.Cls
ReDim warrGMemo(21), mearrGMemo(21)
warrGMemo_NbMax = 20: mearrGMemo_NbMax = 20
blnControl = True
End Sub


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
mearrGFlux(1).Method = "DeleteAll"
'''Call srvgSUB_Update(mearrGFlux(1))
End Sub


Public Sub fgFlux_Display()
Dim I As Integer

fgFlux.Visible = True
fgFlux.Clear: fgFlux_RowDisplay = 0: fgFlux_RowClick = 0
If picGFlux.Height > 700 Then: picGFlux.Cls: Call pic_Resize(picGFlux, 0)

fgFlux.Rows = 1
fgFlux.FormatString = fgFlux_FormatString
fgFlux.Enabled = True
For mearrGFlux_Index = 1 To mearrGFlux_Nb
    meGFlux = mearrGFlux(mearrGFlux_Index)
    fgFlux.Rows = fgFlux.Rows + 1
    fgFlux.Row = fgFlux.Rows - 1
    fgFlux_DisplayLine
Next mearrGFlux_Index

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
For mearrGEch_Index = 1 To mearrGEch_Nb
    meGECh = mearrGEch(mearrGEch_Index)
    fgEch.Rows = fgEch.Rows + 1
    fgEch.Row = fgEch.Rows - 1
    fgEch_DisplayLine
Next mearrGEch_Index

 
fgEch_SortAD = 5
End Sub

Public Sub fgFlux_DisplayLine()
Dim K2 As Integer

paramEffetCommerce.OpérationCode = meGFlux.OpérationCode
srvGSub.param_Opération paramEffetCommerce
fgFlux.Col = 0: fgFlux.Text = GSub_recOpération.Name

If meGFlux.Montant1 <> 0 Then fgFlux.Col = 1: fgFlux.Text = Format(meGFlux.Montant1, "#### ### ###.00 ") & meGFlux.Devise1
If meGFlux.Montant2 <> 0 Then fgFlux.Col = 2: fgFlux.Text = Format(meGFlux.Montant2, "#### ### ###.00 ") & meGFlux.Devise2
fgFlux.Col = 3: fgFlux.Text = dateImp(meGFlux.AmjValeur)
fgFlux.Col = 5: fgFlux.Text = "du " & dateImp(meGFlux.AmjDébut) & " au " & dateImp(meGFlux.AmjFin) & "   (" & meGFlux.Nbj & "j)"
If meGFlux.Taux <> 0 Then fgFlux.Col = 4: fgFlux.Text = Format(meGFlux.Taux, "#0.00000 ") & meGFlux.TauxProvisoire
fgFlux.Col = 6: fgFlux.Text = recStatut_Libellé(meGFlux.Statut & meGFlux.StatutPlus)
fgFlux.Col = 7: fgFlux.Text = meGFlux.IdRéférence & "_" & meGFlux.FluxSéquence
fgFlux.Col = 8: fgFlux.Text = Trim(meGFlux.Application) & "_" & Trim(meGFlux.OpérationCode)
fgFlux.Col = fgFlux.Cols - 1: fgFlux.Text = mearrGFlux_Index

If meGFlux.Statut = "A" Then fgFlux.Col = 0: fgFlux.CellForeColor = errUsr.ForeColor


End Sub

Public Sub fgEch_DisplayLine()
Dim wColor As Long

Select Case meGECh.Statut
    Case " ": wColor = warnUsrColor
    Case "à": wColor = greenColor.ForeColor
    Case Else: wColor = libUsr.ForeColor
End Select

fgEch.Col = 0: fgEch.Text = Trim(meGECh.Application) & "_" & Trim(meGECh.FluxSéquence): fgEch.CellForeColor = wColor
fgEch.Col = 1: fgEch.Text = meGECh.EchFct: fgEch.CellForeColor = wColor
fgEch.Col = 2: fgEch.Text = dateImp(meGECh.EchAMJ) & " - " & timeImp(meGECh.EchHMS): fgEch.CellForeColor = wColor
fgEch.Col = 3: fgEch.Text = meGECh.EchUsr: fgEch.CellForeColor = wColor
fgEch.Col = 4: fgEch.Text = meGECh.ActionFct: fgEch.CellForeColor = wColor
fgEch.Col = 5: fgEch.Text = dateImp(meGECh.ActionAmj) & " - " & timeImp(meGECh.ActionHms): fgEch.CellForeColor = wColor
fgEch.Col = 6: fgEch.Text = meGECh.ActionUsr: fgEch.CellForeColor = wColor
fgEch.Col = 7: fgEch.Text = recStatut_Libellé(meGECh.Statut & meGECh.StatutPlus): fgEch.CellForeColor = wColor
fgEch.Col = 8: fgEch.Text = meGECh.IdRéférence & "_" & meGECh.EchSéquence: fgEch.CellForeColor = wColor
fgEch.Col = fgEch.Cols - 1: fgEch.Text = mearrGEch_Index: fgEch.CellForeColor = wColor

If meGECh.Statut = "A" Then fgEch.Col = 0: fgEch.CellForeColor = errUsr.ForeColor

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
    mearrGFlux_Index = Val(fgFlux.Text)
    fgFlux.Col = fgFlux.Cols - 2
   Select Case lK
        Case 1: fgFlux.Text = Format$(mearrGFlux(mearrGFlux_Index).Montant1, "000000000000000.00")
        Case 2: fgFlux.Text = Format$(mearrGFlux(mearrGFlux_Index).Montant2, "000000000000000.00")
        Case 3: fgFlux.Text = mearrGFlux(mearrGFlux_Index).AmjValeur
        Case 10: fgFlux.Text = Format$(mearrGFlux_Index, "0000000000")
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
    mearrGEch_Index = Val(fgEch.Text)
    fgEch.Col = fgEch.Cols - 2
    Select Case lK
        Case 2: fgEch.Text = mearrGEch(mearrGEch_Index).EchAMJ & mearrGEch(mearrGEch_Index).EchHMS
        Case 5: fgEch.Text = mearrGEch(mearrGEch_Index).ActionAmj & mearrGEch(mearrGEch_Index).ActionHms
        Case 9: fgEch.Text = Format(meGECh.IdRéférence, "000000") & "_" & Format(meGECh.EchSéquence, "0000000")
        Case 11: fgEch.Text = Format$(mearrGEch_Index, "0000000000")
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
For mearrGOpe_Index = 1 To mearrGOpe_Nb
    If mearrGOpe(mearrGOpe_Index).Method <> constIgnore And mearrGOpe(mearrGOpe_Index).Method <> constDelete Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine
    End If
Next mearrGOpe_Index

fgSelect_SortAD = 5
If fgSelect.Rows = 1 Then Exit Sub
'fgSelect_Sort

End Sub
Public Sub fgSelect_DisplayLine()

fgSelect.Col = 0: fgSelect.Text = mearrGOpe(mearrGOpe_Index).RéférenceInterne
fgSelect.Col = 1: fgSelect.Text = mearrGOpe(mearrGOpe_Index).Nature
fgSelect.Col = 2: fgSelect.Text = Format(mearrGOpe(mearrGOpe_Index).Montant1, "#### ### ###.00 ")
fgSelect.Col = 3: fgSelect.Text = mearrGOpe(mearrGOpe_Index).Devise1
fgSelect.Col = 4: fgSelect.Text = dateImp(mearrGOpe(mearrGOpe_Index).AmjDébut)
fgSelect.Col = 5: fgSelect.Text = dateImp(mearrGOpe(mearrGOpe_Index).AmjFin)
fgSelect.Col = 6: fgSelect.Text = Compte_Imp(mearrGOpe(mearrGOpe_Index).EngagementCompte)
Call CV_AttributS(mearrGOpe(mearrGOpe_Index).Devise1, meCV1)
recCompteInit recCompte
recCompte.Société = SocId$
recCompte.Agence = SocAgence$
recCompte.Devise = meCV1.DeviseN
recCompte.Numéro = mearrGOpe(mearrGOpe_Index).EngagementCompte
mdbCptP0_Find recCompte
fgSelect.Col = 7: fgSelect.Text = recCompte.Intitulé
fgSelect.Col = 8: fgSelect.Text = mearrGOpe(mearrGOpe_Index).RéférenceExterne
fgSelect.Col = 9: fgSelect.Text = recStatut_Libellé(mearrGOpe(mearrGOpe_Index).Statut & mearrGOpe(mearrGOpe_Index).StatutPlus)
fgSelect.Col = 10: fgSelect.Text = mearrGOpe(mearrGOpe_Index).IdRéférence
fgSelect.Col = fgSelect_arrIndex - 1: fgSelect.Text = ""
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = mearrGOpe_Index
If mearrGOpe(mearrGOpe_Index).Statut = "à" Then
    For I = 0 To fgSelect_arrIndex
      fgSelect.Col = I: fgSelect.CellForeColor = warnUsrColor
    Next I
End If

End Sub
Public Sub fgSelect_Load()
Dim X As String, mMethod As String

recGOpe_Init xGOpe
xGOpe.Application = paramEffetCommerce.Application

Select Case currentAction
    Case constDemandeDeValidation
            xGOpe.Method = "SnapLA"
            xGOpe.IdRéférence = 0
            xGOpe.Statut = "à"
            xGOpe.StatutPlus = "B "
            mearrGOpe(0) = xGOpe
            mearrGOpe(0).IdRéférence = 999999999#
    Case constValider
            xGOpe.Method = "SnapLA"
            xGOpe.IdRéférence = 0
            xGOpe.Statut = "à"
             xGOpe.StatutPlus = "V "
           
            mearrGOpe(0) = xGOpe
            mearrGOpe(0).IdRéférence = 999999999#

    Case "mnuList"
            xGOpe.Method = "SnapLRI"
            X = Trim(txtSelect)
            xGOpe.RéférenceInterne = X
            xGOpe.IdRéférence = 0
            xGOpe.Statut = " "
            
            mearrGOpe(0) = xGOpe
            mearrGOpe(0).IdRéférence = 999999999#
            mearrGOpe(0).RéférenceInterne = X & "9z"

End Select

Call srvGOpe_Load(xGOpe, mearrGOpe(0))

mearrGOpe_Nb = srvGOpe.arrGOpe_NB
mearrGOpe_NbMax = mearrGOpe_Nb + 1: ReDim mearrGOpe(mearrGOpe_NbMax)

For I = 1 To mearrGOpe_Nb
    mearrGOpe(I) = srvGOpe.arrGOpe(I)
Next I
If mearrGOpe_Nb = 0 Then Call lstErr_Clear(lstErr, cmdContext, "? aucune opération sélectionnée")

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
    mearrGOpe_Index = Val(fgSelect.Text)
    fgSelect.Col = fgSelect_arrIndex - 1
    Select Case lK
        Case 1: fgSelect.Text = mearrGOpe(mearrGOpe_Index).Nature & Trim(mearrGOpe(mearrGOpe_Index).RéférenceInterne)
        Case 2: fgSelect.Text = Format$(mearrGOpe(mearrGOpe_Index).Montant1, "000000000000000.00") & mearrGOpe(mearrGOpe_Index).Devise1
        Case 3: fgSelect.Text = mearrGOpe(mearrGOpe_Index).Devise1 & Format$(mearrGOpe(mearrGOpe_Index).Montant1, "000000000000000.00")
        Case 4: fgSelect.Text = mearrGOpe(mearrGOpe_Index).AmjDébut & Trim(mearrGOpe(mearrGOpe_Index).RéférenceInterne)
        Case 5: fgSelect.Text = mearrGOpe(mearrGOpe_Index).AmjFin & Trim(mearrGOpe(mearrGOpe_Index).RéférenceInterne)
        Case 900: fgSelect.Text = mearrGOpe(mearrGOpe_Index).EngagementCompte & mearrGOpe(mearrGOpe_Index).Nature & mearrGOpe(mearrGOpe_Index).Devise1
        Case 910: fgSelect.Text = mearrGOpe(mearrGOpe_Index).RéférenceExterne
        Case fgSelect_arrIndex: fgSelect.Text = Format$(mearrGOpe_Index, "0000000000")
    End Select
Next I

fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub


Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents
If Not IsNull(srvGSub.param_Init(paramEffetCommerce, cboNature)) Then MsgBox "jpl: unload " 'Unload Me
If Not IsNull(srvGSub_EC.param_Init(paramEffetCommerce)) Then MsgBox "jpl: unload " 'Unload Me
cboNature.AddItem " "


' Chargement des devise autorisées pour les opérations de TC
recElpTable_Init recElpTable
recElpTable.Id = paramEffetCommerce.TableId
recElpTable.K1 = "Devise"
Call cbo_Load(recElpTable, cboDevise1, 3)
meCV1 = CV_Euro: meCV2 = CV_Euro: meCV3 = CV_Euro


SSTab1.Tab = 0
tableElpTable_Open
'paramAmjEngagementMin = paramAmjOpérationMin
paramAmjEngagementMax = dateElp("Ouvré", 7, DSys)
ReDim mearrGOpe(10)
ReDim mearrGEch(10): mearrGEch_NbMax = 10
ReDim mearrGFlux(10): mearrGFlux_NbMax = 10

cmdReset

mLotàComptabiliserValider = mnuLotàComptabiliserValider.Caption
mLotàComptabiliserAnnuler = mnuLotàComptabiliserAnnuler.Caption
mLotàComptabiliserPrint = mnuLotàComptabiliserPrint.Caption

mnuOpérationSaisir.Enabled = EffetCommerceAut.Saisir
mnuListàValider.Enabled = EffetCommerceAut.Consulter
mnuList.Enabled = EffetCommerceAut.Consulter
mnuValidationDemande.Enabled = EffetCommerceAut.Saisir
mnuComptaEchéancier.Enabled = EffetCommerceAut.Consulter
mnuListEchéancier.Enabled = EffetCommerceAut.Consulter
mnuComptaLotsàValider.Enabled = EffetCommerceAut.Comptabiliser
mnuComptaLotComptabiliséAnnuler.Enabled = EffetCommerceAut.Xspécial
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


Private Sub cboDevise1_Click()
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

Private Sub chkComEndos_Click()
If blnControl Then cmdControl

End Sub

Private Sub chkComManipulation_Click()
If blnControl Then cmdControl

End Sub


Private Sub chkComptaReprise_Click()
If blnControl Then cmdControl

End Sub


Private Sub chkTauxMajoré_Click()
If blnControl Then cmdControl

End Sub


Private Sub chkTauxNonAccepté_Click()
If blnControl Then cmdControl

End Sub


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdOk_Click()
Dim V

Me.Enabled = False

cmdControl
If lstErr.ListCount <> 0 Then GoTo Exit_Sub
Select Case cmdOk.Caption
    Case constàValider
        mearrGEch_Nb = 1: srvGSub.Gech_Update mearrGEch(1)
        mearrGFlux_Nb = 0
        meGOpe.Statut = "à"
        meGOpe.StatutPlus = "B "
        
        blnMsgTxt_Concat_Transaction = False ''' !!! initialisation d'IdRéférence à la création du dossier et propagation GECH et GMEMO
        V = srvGOpe_Update(meGOpe)
        meGOpe.Method = ""
        blnMsgTxt_Concat_Transaction = blnMsgTxt_Concat

    Case constValider
        If Not EffetCommerceAut.Xspécial And Trim(arrGECh(1).EchUsr) = Trim(usrId) Then
            Call MsgBox("Vous ne pouvez pas valider vos propres opérations.", vbCritical, "TC : Validation ")
            Call lstErr_AddItem(lstErr, cmdContext, "? validation interdite")
        Else
            cmdValidation_Ok
        End If
Case Else
        Call lstErr_AddItem(lstErr, cmdContext, "? cmdOk : " & cmdOk.Caption)
End Select

If lstErr.ListCount = 0 Then V = cmdSave_Db

If IsNull(V) Then currentAction = "cmdOk": cmdContext_Quit

Exit_Sub:

Me.Enabled = True
AppActivate Me.Caption

fraOpération_Load " "
End Sub

Private Sub cmdOpérationElpDisplay_Click()
srvGOpe_ElpDisplay meGOpe
End Sub

Private Sub cmdPrint_Click()
If SSTab1.Tab > 1 Then
    cmdPrint_Dossier
End If

End Sub

Private Sub cmdSave_Click()

cmdControl
lstErr.Clear
frmTC.Enabled = False
Select Case cmdSave.Caption
    Case constEnAttente
        mearrGEch_Nb = 1: srvGSub.Gech_Update mearrGEch(1)
        mearrGFlux_Nb = 0
       meGOpe.Statut = "à"
        meGOpe.StatutPlus = "? "
        cmdSave_Db
        cmdContext_Quit
    Case constàModifier
        mearrGEch_Nb = 0
        mearrGFlux_Nb = 0
        meGOpe.Statut = "à"
        meGOpe.StatutPlus = "? "
        cmdSave_Db
        cmdContext_Quit
    Case constEffacer
     '   meGOpe.Method = constDelete
    Case Else
        Call lstErr_AddItem(lstErr, cmdContext, "? cmdsave : " & cmdSave.Caption)
End Select

''If lstErr.ListCount = 0 Then cmdSave_Db
frmTC.Enabled = True

End Sub

Private Sub cmdSelect_Click()
'If optSelectDemandeValidation Then cmdDemandeValidation
'If optSelectValidation Then cmdValidation_Load

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
        mearrGEch_Index = Val(fgEch.Text)
        meGECh = mearrGEch(mearrGEch_Index)
        Call fgEch_Color(fgEch_RowClick, MouseMoveUsr.BackColor, fgEch_ColorClick)
        Call srvGSub.pic_Resize(picGEch, 0)
        If Button = vbRightButton Then
            mnuGEchElpDisplay.Enabled = EffetCommerceAut.Xspécial
            If meGECh.FluxSéquence > 0 Then mnuGEchAction.Enabled = True
            Me.PopupMenu mnuGEch, vbPopupMenuLeftButton
        Else
            If meGECh.FluxSéquence > 0 Then mnuGEchAction_Click
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
        mearrGFlux_Index = Val(fgFlux.Text)
        '''meGFlux.CptMvtLot = mearrGFlux(mearrGFlux_Index).CptMvtLot
        meGFlux = mearrGFlux(mearrGFlux_Index)
        Call fgFlux_Color(fgFlux_RowClick, MouseMoveUsr.BackColor, fgFlux_ColorClick)
        Call srvGSub.pic_Resize(picGFlux, 0)
       
         If Button = vbRightButton Then
            mnuGFluxElpDisplay.Enabled = EffetCommerceAut.Xspécial
            mnuGFluxGEch.Enabled = True
            mnuGFluxAction.Enabled = True
           Me.PopupMenu mnuGFlux, vbPopupMenuLeftButton
        Else
            mnuGFluxAction_Click
        End If
        
        If currentAction = constDisplay Then
'$$$            Param_OpérationCode meGFlux.OpérationCode
'$$$            mnuEchéancier_Set
'$$$            Me.PopupMenu mnuEchéancier, vbPopupMenuLeftButton
           Else
'$$$            If meGFlux.CptMvtLot > 0 Then
'$$$                mnuLotàComptaValidation = False
'$$$                mnuLotàComptaAnnulation = False
'$$$                mnuLotàComptaAnnulation = False
              
 '$$$               If meGFlux.Statut = "à" And meGFlux.StatutPlus = "C " Then
 '$$$                   mnuLotàComptaValidation = EffetCommerceAut.Comptabiliser
  '$$$                  mnuLotàComptaAnnulation = EffetCommerceAut.Comptabiliser
  '$$$                  mnuLotàComptaPrint = EffetCommerceAut.Comptabiliser
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
        mearrGOpe_Index = Val(fgSelect.Text)
        mGOpe = mearrGOpe(mearrGOpe_Index)
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    
        If mGOpe.IdRéférence > 0 Then
            Select Case currentAction
                Case constValider: fgSelect_MouseDown_Validation
                Case Else: fgSelect_MouseDown_Opération
            End Select
        End If
    End If
End If

End Sub
Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Private Sub mnuComptaLotsàValider_Click()
Dim I As Integer, mRupture As String, x6 As String
Me.Enabled = False
libRéférenceInterne = "Validation des bordereaux d'effets"

lstErr.Clear
currentAction = constValider
fgSelect_Load

If lstErr.ListCount <> 0 Then GoTo Exit_Sub

fgSelect_SortX 910   ' tri Référence externe

Exit_Sub:

Me.Enabled = True
AppActivate Me.Caption

End Sub

Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub

Private Sub mnuContextQuitter_Click()
Unload Me
End Sub

Private Sub mnuGEchAction_Click()
If Trim(meGECh.ActionFct) <> "" Then
    V = GMemo_Scan(meGECh.EchSéquence)
    Call GMemo_Display(picGEch, warrGMemo_Nb, warrGMemo())
Else
    V = GFlux_Scan(meGECh.FluxSéquence)
    If IsNull(V) Then
        Call srvGSub_EC.GMemo_Gen(paramEffetCommerce, meGOpe, meGFlux, meGECh, warrGMemo_Nb, warrGMemo())
        Call GMemo_Display(picGEch, warrGMemo_Nb, warrGMemo())
    End If
End If
End Sub

Private Sub mnuGEchElpDisplay_Click()
srvGEch_ElpDisplay meGECh


End Sub

Private Sub mnuGFluxAction_Click()
Dim I As Integer, J As Integer

GEch_Scan meGFlux.FluxSéquence

xarrGMemo_Nb = 0:
For I = 1 To warrGEch_Nb
    meGECh = warrGEch(I)
    If Trim(meGECh.ActionFct) <> "" Then
        V = GMemo_Scan(meGECh.EchSéquence)
    Else
        V = GFlux_Scan(meGECh.FluxSéquence)
        If IsNull(V) Then
            Call srvGSub_EC.GMemo_Gen(paramEffetCommerce, meGOpe, meGFlux, meGECh, warrGMemo_Nb, warrGMemo())
        End If
    End If
    ReDim Preserve xarrGMemo(xarrGMemo_Nb + warrGMemo_Nb + 1)
    For J = 1 To warrGMemo_Nb
        xarrGMemo_Nb = xarrGMemo_Nb + 1
        xarrGMemo(xarrGMemo_Nb) = warrGMemo(J)
    Next J
Next I
Call GMemo_Display(picGFlux, xarrGMemo_Nb, xarrGMemo())
End Sub

Private Sub mnuGFluxElpDisplay_Click()
srvGFlux_ElpDisplay meGFlux

End Sub

Private Sub mnuGFluxGEch_Click()
GEch_Scan meGFlux.FluxSéquence
Call srvGSub.GEch_Display(picGFlux, warrGEch_Nb, warrGEch())

End Sub

Private Sub mnuList_Click()
currentAction = "mnuList"
fgSelect_Load

End Sub

Private Sub mnuListàValider_Click()
currentAction = "mnuListàValider"
fgSelect_Load

End Sub

Private Sub mnuLotàComptabiliserValider_Click()
Dim mRupture As String
MsgBox "Vérifier Saisisseur / Valideur"
'If Not EffetCommerceAut.Xspécial And Trim(mearrGEch(1).EchUsr) = Trim(usrId) Then
'      Call lstErr_Clear(lstErr, cmdContext, "! Vous ne pouvez pas valider vos opérations")
'  Else
'      mnuOpérationValider = EffetCommerceAut.Valider
'  End If
Me.Enabled = False
 mRupture = mGOpe.RéférenceExterne
Call lstErr_Clear(lstErr, lstErr, mRupture & " Début")
Call lstErr_AddItem(lstErr, lstErr, " ")

For I = 1 To mearrGOpe_Nb
    If mRupture = mearrGOpe(I).RéférenceExterne Then
        Call lstErr_ChangeLastItem(lstErr, lstErr, mGOpe.IdRéférence & " GEN")
        mGOpe = mearrGOpe(I)
        meGOpe = mGOpe
        GEch_Load
        If mearrGEch_Nb = 1 Then: ReDim Preserve mearrGEch(10): Call GEch_Gen(paramEffetCommerce, meGOpe, mearrGEch_Nb, mearrGEch())
        Call GFlux_Gen(paramEffetCommerce, meGOpe, mearrGFlux_Nb, mearrGFlux())
        
        cmdValidation_Ok
        Call lstErr_ChangeLastItem(lstErr, lstErr, mGOpe.IdRéférence & " MAJ")
    V = cmdSave_Db
  End If

Next I

Call lstErr_AddItem(lstErr, lstErr, mRupture & "Fin")

Exit_Sub:

Me.Enabled = True
AppActivate Me.Caption

End Sub

Private Sub mnuOpérationDisplay_Click()
fraOpération_Load " "
End Sub

Private Sub mnuOpérationModifier_Click()
If EffetCommerceAut.Saisir Then
    mGOpe.Method = constUpdate
    mnuOpérationSaisir_Init
    fraOpération_Load "Update"
    cmdSave.Visible = True
    blncmdOk_Visible = True: blncmdSave_Visible = True
    currentAction = constSaisie
    fraOpération1.Enabled = True
    fraOpération2.Enabled = True
    blnSetfocus = True: currentActiveControl_Name = "txtMontant1"
    blnControl = True: cmdControl
End If

End Sub

Private Sub mnuOpérationSaisir_Click()
If EffetCommerceAut.Saisir Then
    recGOpe_Init mGOpe

    mGOpe.Method = constAddNew
    currentAction = constSaisie
    mnuOpérationSaisir_Init
    meGOpe = mGOpe
End If

End Sub


Private Sub mnuOpérationValider_Click()
If EffetCommerceAut.Valider Then
    mGOpe.Method = constUpdate
    mnuOpérationSaisir_Init
    fraOpération_Load "Update"
    blncmdOk_Visible = True: blncmdSave_Visible = True
    cmdOk.Visible = True
    cmdSave.Visible = True
    currentAction = constValider
    cmdContext.Caption = constcmdAbandonner
    fraOpération1.Enabled = True: fraOpération2.Enabled = True
End If
End Sub

Private Sub mnuValidationDemande_Click()
Dim I As Integer, mRupture As String, x6 As String
Me.Enabled = False
libRéférenceInterne = constDemandeDeValidation
lstErr.Clear
currentAction = constDemandeDeValidation
fgSelect_Load

If lstErr.ListCount <> 0 Then GoTo Exit_Sub

fgSelect_SortX 900   ' tri compte Nature Devise ??? Date d'échéance à prèciser
mRupture = "???"

For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex - 1
    X = fgSelect.Text
    If mRupture <> X Then
        mRupture = X
        x6 = InputBox("N° bordereau")
        x6 = Format$(x6, "000000")
    End If
    
    fgSelect.Col = fgSelect_arrIndex
    mearrGOpe_Index = Val(fgSelect.Text)
    mearrGOpe(mearrGOpe_Index).Method = constUpdate
    mearrGOpe(mearrGOpe_Index).StatutPlus = "V "
    Mid$(mearrGOpe(mearrGOpe_Index).RéférenceExterne, 11, 6) = x6
        blnMsgTxt_Concat_Transaction = False ''' !!! à revoir pour maj par bloc
    srvGOpe_Update mearrGOpe(mearrGOpe_Index)
        blnMsgTxt_Concat_Transaction = blnMsgTxt_Concat

Next I

mRupture = "???"
prtEffetCommerce_Avis_Open

''Call MsgBox("effet Commerce", vbInformation, "pas impr de bordereau")
''GoTo Exit_Sub

For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    mearrGOpe_Index = Val(fgSelect.Text)
    If mRupture <> mearrGOpe(mearrGOpe_Index).RéférenceExterne Then
        mRupture = mearrGOpe(mearrGOpe_Index).RéférenceExterne
        prtEffetCommerce_Avis_Form mearrGOpe(mearrGOpe_Index)
    End If

    prtEffetCommerce_Avis_Line mearrGOpe(mearrGOpe_Index)
Next I

prtEffetCommerce_Close

Exit_Sub:

Me.Enabled = True
AppActivate Me.Caption

End Sub

Private Sub txtAmjEchéance_Change()
If blnControl Then cmdControl

End Sub

Private Sub txtAmjEchéance_GotFocus()
DTPicker_GotFocus txtAMJEchéance


End Sub


Private Sub txtAMJEchéance_LostFocus()
DTPicker_LostFocus txtAMJEchéance
If blnControl Then cmdControl

End Sub


Private Sub txtAmjEngagement_Change()
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

Private Sub txtRéférenceExterne_GotFocus()
txt_GotFocus txtRéférenceExterne

End Sub

Private Sub txtRéférenceExterne_LostFocus()
txt_LostFocus txtRéférenceExterne
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

paramEffetCommerce.TableId = "GFlux_EC"
Call BiaPgmAut_Init(mId$(Msg, 1, 12), EffetCommerceAut)    ' "SOBF_Effets"

Form_Init

End Sub


Public Sub cmdControl_OpérationG()

If mtxtEngagementCompte = txtEngagementCompte And mcboNature_ListIndex = cboNature.ListIndex Then Exit Sub
meGOpe.AmjDébut = DSys
meGOpe.AmjFin = DSys


Call cbo_Value(meGOpe.Nature, cboNature)

paramEffetCommerce.NatureCode = meGOpe.Nature
V = srvGSub.param_Nature(paramEffetCommerce)
If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, Trim(V) & " nature")

libEngagementCompte = ""
If Trim(txtEngagementCompte) = "" Then
    meRacine.Numéro = 0
    Call lstErr_AddItem(lstErr, lstErr, "? préciser la contrepartie")
    If Not blnSetfocus Then blnSetfocus = True: currentActiveControl_Name = "txtEngagementCompte"

Else
    meRacine.Numéro = CLng(num_CDec(mId$(Trim(txtEngagementCompte), 1, 5)))
    V = srvRacineFind(meRacine)
''   MsgBox "cmdcontrol_opérationg", vbInformation, "Racine ok"
If blnJPL Then V = Null
    If Not IsNull(V) Then
        Call lstErr_AddItem(lstErr, cmdContext, "? contrepartie inconnue")
    Else
        txtEngagementCompte = meRacine.Numéro
        libEngagementCompte = meRacine.Intitulé
    End If
End If
If lstErr.ListCount = 0 And currentAction = constSaisie Then
    mtxtEngagementCompte = txtEngagementCompte
    mcboNature_ListIndex = cboNature.ListIndex
    meCV1.DeviseIso = "": meCV2.DeviseIso = ""
    meGOpe.EngagementCompte = "": meGOpe.EchéanceCompte = ""
    blnSetfocus = True: currentActiveControl_Name = "txtMontant1"
    mGOpe = meGOpe
    If meGOpe.Nature = "LCEnc" Then
        fraConditions_Escompte.Enabled = False
    Else
        fraConditions_Escompte.Enabled = True
    End If
    If meGOpe.Nature = "MCNE" Then
        txtAmjMCNE.Enabled = False
    Else
        txtAmjMCNE.Enabled = True
    End If
    
Else
    mcboNature_ListIndex = -2
End If


End Sub

Public Sub cmdContext_Return()
If fraOption.Visible Then
    fraOption.Visible = False
Else
    If SSTab1.Tab = 0 Then
        'mnuList_Click
    Else
        If currentActiveControl_Name = "txtRéférenceInterne" Then
            SSTab1.Tab = 3
            If txtTaux.Enabled = True Then txtTaux.SetFocus
        Else
            SendKeys "{TAB}"
        End If
        
    End If
End If

End Sub



Public Sub cmdControl_OpérationD()
Dim I1 As Integer, I2 As Integer, X As String, mAmjEscMin As String * 8
Dim wCur As Currency, wL As Long

meGOpe.optReprise = chkComptaReprise

Call cbo_Value(meGOpe.Devise1, cboDevise1)
If meGOpe.Devise1 <> xGOpe.Devise1 Then cmdControl_OpérationD_Devise
meGOpe.Devise2 = meGOpe.Devise1

If currentAction = constSaisie Then
    meGOpe.RéférenceExterne = Format$(meRacine.Numéro, "00000") & meGOpe.Nature

    meGOpe.EngagementCompte = Format$(meRacine.Numéro, "00000") & paramEffetCommerce.BiatypEngagement & "010"
    Compte_BiaClé meGOpe.EngagementCompte

    If meGOpe.EngagementCompte <> xGOpe.EngagementCompte Then
        libCompte1 = Compte_Imp(meGOpe.EngagementCompte)
        V = Compte_Load(meCV1.DeviseN, meGOpe.EngagementCompte)
        If Not IsNull(V) Then
            Call lstErr_AddItem(lstErr, lstErr, "? " & V):
            meGOpe.EngagementCompte = ""
            libCompte1.ForeColor = errUsr.ForeColor
        Else
            libCompte1.ForeColor = libUsr.ForeColor
        End If
    
    End If

        meGOpe.EchéanceCompte = Format$(meRacine.Numéro, "00000") & paramEffetCommerce.BiatypEchéance & "010"
        Compte_BiaClé meGOpe.EchéanceCompte
    If meGOpe.EchéanceCompte <> xGOpe.EchéanceCompte Or Trim(meGOpe.EchéanceCompte) = "" Then
        libCompte2 = Compte_Imp(meGOpe.EchéanceCompte)
        V = Compte_Load(meCV2.DeviseN, meGOpe.EchéanceCompte)
        If Not IsNull(V) Then
            Call lstErr_AddItem(lstErr, lstErr, "? " & V):
            meGOpe.EchéanceCompte = ""
            libCompte2.ForeColor = errUsr.ForeColor
        Else
            libCompte2.ForeColor = libUsr.ForeColor
        End If
    End If
End If

meGOpe.EngagementCorrCompte = paramEffetCommerce_CompteRecouvreur
meGOpe.EchéanceCorrCompte = paramEffetCommerce_CompteCompensateur

If Trim(txtMontant1) = "" Then
    Call lstErr_AddItem(lstErr, txtMontant1, " ? préciser le montant")
    If Not blnSetfocus Then blnSetfocus = True: currentActiveControl_Name = "txtMontant1"
    Exit Sub
Else
    
   meGOpe.Montant1 = CCur(num_CDec(txtMontant1))
End If

If lstErr.ListCount <> 0 Then GoTo Exit_Sub

Call DTPicker_Control(txtAmjEngagement, meGOpe.AmjDébut)
If currentAction = constSaisie Or currentAction = constValider Then
    If meGOpe.AmjDébut < paramEffetCommerce_AmjRemMax Then Call lstErr_AddItem(lstErr, txtAmjEngagement, "? date de remise < " & dateImp(paramEffetCommerce_AmjRemMax))
    If meGOpe.AmjDébut > DSys Then Call lstErr_AddItem(lstErr, txtAmjEngagement, "? date de remise > " & dateImp(DSys))
End If

Call DTPicker_Control(txtAMJEchéance, meGOpe.AmjFin)
If meGOpe.AmjDébut > meGOpe.AmjFin Then Call lstErr_AddItem(lstErr, txtAMJEchéance, "? date d'engagement < date d'échéance ")

meGOpe.AmjEngagement = DSys

If meGOpe.Nature = "LCEnc" Then
    X = DateElp_X(paramEffetCommerce_NbjEncMin, DSys)
    If meGOpe.AmjFin > X Then X = meGOpe.AmjFin
    meGOpe.AmjEchéance1 = DateElp_X(paramEffetCommerce_NbjEncCpt, X)
Else
    meGOpe.AmjEchéance1 = meGOpe.AmjFin
End If


X = Trim(txtTiréNom)
If X = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le nom du tiré")
Mid$(mearrGMemo(1).MemoText, 1, 50) = X

X = Trim(txtTiréDomiciliation)
If X = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser la domiciliation du tiré")
Mid$(mearrGMemo(1).MemoText, 51, 50) = X

X = Trim(txtTiréRéférence)
If X = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser la Référence de la facture")
Mid$(mearrGMemo(1).MemoText, 101, 50) = X

Call DTPicker_Control(txtAmjMCNE, X)
Mid$(mearrGMemo(1).MemoText, 151, 8) = X


X = Trim(txtRéférenceInterne)
If X = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser la référence de l'effet")
meGOpe.RéférenceInterne = X

If lstErr.ListCount <> 0 Then GoTo Exit_Sub

wTaux = 0

If meGOpe.Nature <> "LCEnc" Then

    If Trim(txtTaux) = "" Then
        meGOpe.TauxMarge1 = 0
        meGOpe.TauxMarge2 = 0
        Call lstErr_AddItem(lstErr, lstErr, " ? préciser le cours")
        If Not blnSetfocus Then blnSetfocus = True: currentActiveControl_Name = "txtTaux"
    Else
        meGOpe.TauxMarge2 = Round(CDbl(num_CDec(txtTaux)), 5)
        wTaux = meGOpe.TauxMarge2
        If chkTauxMajoré = "1" Then wTaux = wTaux + Round(CDbl(num_CDec(txtTauxMajoré)), 5)
        If chkTauxNonAccepté = "1" Then wTaux = wTaux + Round(CDbl(num_CDec(txtTauxNonAccepté)), 5)
        meGOpe.TauxMarge1 = wTaux
 ' nbj minimun d'escompte
        xGOpe = meGOpe
        mAmjEscMin = DateElp_X(paramEffetCommerce_NbjEscMin, meGOpe.AmjDébut)
        If meGOpe.AmjFin < mAmjEscMin Then meGOpe.AmjFin = mAmjEscMin

        V = fctGOpe_Intérêts(meGOpe, meCV1, wCur, wL)
        meGOpe.AmjDébut = xGOpe.AmjDébut
        meGOpe.PériodeNb = wL
        meGOpe.Mensualité = wCur
        If wL <= 90 Then
            If chkTauxMajoré = "1" Then Call lstErr_AddItem(lstErr, cmdContext, "? taux majoré pour " & wL & " jours ?")
        Else
            If Not blnTauxMajoré_Set Then
                blnTauxMajoré_Set = True: chkTauxMajoré = 1
            Else
                If chkTauxMajoré <> "1" Then
                    If Not blnTauxMajoré_MsgBox Then blnTauxMajoré_MsgBox = True: Call MsgBox("Vous avez déchoché la case 'taux majoré > 90 jours'", vbExclamation, "Effet à l'escompte")
                End If
            End If
            
        End If
        
    End If
    
End If


Mid$(meGOpe.TauxRéférence2, 3, 1) = chkTauxMajoré

Mid$(meGOpe.TauxRéférence2, 4, 1) = chkTauxNonAccepté




Mid$(meGOpe.TauxRéférence2, 1, 1) = chkComEndos
If chkComEndos = "1" Then
    meGOpe.Frais1 = Round(meGOpe.Montant1 * Round(CDbl(num_CDec(txtComEndos)) / 1000, 5), 2)
Else
    meGOpe.Frais1 = 0
End If


Mid$(meGOpe.TauxRéférence2, 2, 1) = chkComManipulation
If chkComManipulation = "1" Then
    meGOpe.Frais2 = Round(CDbl(num_CDec(txtComManipulation)), 2)
Else
    meGOpe.Frais2 = 0
End If

meGOpe.Frais3 = Round(meGOpe.Frais2 * tauxTVA, 2)

meGOpe.Montant2 = meGOpe.Mensualité + meGOpe.Frais1 + meGOpe.Frais2 + meGOpe.Frais3


Exit_Sub:

fraOpération_Display

End Sub
Public Sub fraOpération_Display()

If meGOpe.Montant1 = 0 Then
    txtMontant1 = ""
Else
    txtMontant1 = Format$(meGOpe.Montant1, "### ### ### ###.00")
End If
txtRéférenceExterne = Trim(meGOpe.RéférenceExterne)
txtRéférenceInterne = Trim(meGOpe.RéférenceInterne)

If meGOpe.TauxMarge2 = 0 Then
    txtTaux = ""
Else
    txtTaux = Trim(Format$(meGOpe.TauxMarge2, "### ##0.00000"))
End If
libIntérêts = Format$(meGOpe.Mensualité, "### ### ##0.00")
libFrais1 = Format$(meGOpe.Frais1, "### ### ##0.00")
libFrais2 = Format$(meGOpe.Frais2, "### ### ##0.00")
libFrais3 = Format$(meGOpe.Frais3, "### ### ##0.00")

libInfo = Trim(cboNature.List(cboNature.ListIndex)) & Chr$(13) _
        & "Brut    : " & Format$(meGOpe.Montant1, "### ### ### ##0.00") & " " & meGOpe.Devise2 & Chr$(13) _
        & " Net     : " & Format$(meGOpe.Montant1 - meGOpe.Montant2, "### ### ### ##0.00") & Chr$(13) _
        & meRacine.Numéro & "   : " & meRacine.Intitulé
libStatut = "Statut         : " & recStatut_Libellé(mGOpe.Statut & mGOpe.StatutPlus) & Chr$(13) _
            & "Référence  : " & Format$(mGOpe.IdRéférence, "#### ### ##0") & Chr$(13)

End Sub

Public Function cmdSave_Db()
''If lstErr.ListCount = 0 Then


If blnMsgTxt_Concat_Transaction Then sndMsgTxt_Init
    blnControl = False
    
    If Trim(meGOpe.Method) <> "" Then V = srvGOpe_Update(meGOpe)
    
    xGOpe = meGOpe
        
    If IsNull(V) Then V = GEch_Save(meGOpe.IdRéférence)
    If IsNull(V) Then V = GFlux_Save(meGOpe.IdRéférence)
    If IsNull(V) Then V = GMemo_Save(meGOpe.IdRéférence)

    If IsNull(V) And blnComptaAuto Then
        mearrGMemo(1).Method = constCompta
        mearrGMemo(1).MemoSéquence = 0
        V = srvGMemo_Update(mearrGMemo(1))
    End If
If blnMsgTxt_Concat_Transaction Then V = sndMsgTxt_Ok

cmdSave_Db = V

    If IsNull(V) Then
        If blnfgSelect_DisplayLine Then
            mearrGOpe(mearrGOpe_Index) = meGOpe
            If meGOpe.Method = constDelete Then
                fgSelect_Display
            Else
                fgSelect_DisplayLine
            End If
        End If
        lastActiveControl_Name = ""
        cmdOk.Visible = False
        cmdSave.Visible = False
        Call lstErr_Clear(lstErr, cmdContext, "Mise à jour effectuée - Identification : " & meGOpe.IdRéférence)
    ''    cmdContext_Quit
    Else
        Call lstErr_Clear(lstErr, cmdContext, V)
 ''''       cmdReset
    End If

''End If

End Function




Public Sub mnuOpérationSaisir_Init()
cmdReset
blnControl = False
SSTab1.Tab = 2
fgSelect.Enabled = False
fgFlux.Enabled = False
fgFlux.Clear: fgFlux.Rows = 1: fgFlux_RowDisplay = 0
fraOpération1.Enabled = True
fraOpération2.Enabled = True

blncmdOk_Visible = True: blncmdSave_Visible = True
blnAmjEchéance = False
txtRéférenceExterne.Enabled = False
cboNature.SetFocus: 'txtEngagementCompte.SetFocus
cmdContext.Caption = constcmdAbandonner

chkTauxMajoré.Value = "0": txtTauxMajoré.Enabled = False: txtTauxMajoré = Format$(paramEffetCommerce_TauxMargeMajoré, "#0.00")
chkTauxNonAccepté.Value = "0": txtTauxNonAccepté.Enabled = False: txtTauxNonAccepté = Format$(paramEffetCommerce_TauxMargeNonAccepté, "#0.00")
chkComManipulation.Value = "1": txtComManipulation.Enabled = False: txtComManipulation = ""
chkComEndos.Value = "1": txtComEndos.Enabled = False: txtComEndos = ""
blnTauxMajoré_Set = False: blnTauxMajoré_MsgBox = False

Call srvGSub.GEch_New(paramEffetCommerce, meGOpe, mearrGEch_Nb, mearrGEch())
Call srvGSub.GMemo_New(paramEffetCommerce, meGOpe, mearrGMemo_Nb, mearrGMemo())

blnControl = True
End Sub


Public Function GFlux_Scan(lFluxSéquence As Long)
GFlux_Scan = "?GFlux_Scan"
For mearrGFlux_Index = 1 To mearrGFlux_Nb
    If mearrGFlux(mearrGFlux_Index).FluxSéquence = lFluxSéquence Then
        meGFlux = mearrGFlux(mearrGFlux_Index)
        GFlux_Scan = Null
        Exit For
    End If
Next mearrGFlux_Index

End Function
Public Function GMemo_Scan(lEchSéquence As Long)
GMemo_Scan = "?GMemo_Scan"
warrGMemo_Nb = 0
For mearrGMemo_Index = 1 To mearrGMemo_Nb
    If mearrGMemo(mearrGMemo_Index).EchSéquence = lEchSéquence Then
        If warrGMemo_Nb = warrGMemo_NbMax Then warrGMemo_NbMax = warrGMemo_NbMax + 10: ReDim Preserve warrGMemo(warrGMemo_NbMax)
        warrGMemo_Nb = warrGMemo_Nb + 1
        warrGMemo(warrGMemo_Nb) = mearrGMemo(mearrGMemo_Index)
    End If
Next mearrGMemo_Index
End Function


Public Function GEch_Scan(lFluxSéquence As Long)
GEch_Scan = "?GEch_Scan"
warrGEch_Nb = 0
ReDim warrGEch(mearrGEch_Nb + 1)
For mearrGEch_Index = 1 To mearrGEch_Nb
    If mearrGEch(mearrGEch_Index).FluxSéquence = lFluxSéquence Then
        warrGEch_Nb = warrGEch_Nb + 1
        warrGEch(warrGEch_Nb) = mearrGEch(mearrGEch_Index)
    End If
Next mearrGEch_Index
End Function

Public Sub cmdValidation_Ok()
Dim I As Integer, K As Integer

meGOpe.Method = constUpdate
meGOpe.Statut = " "
meGOpe.StatutPlus = "  "

With mearrGEch(1)                                  ' Saisie
    .Method = constUpdate
    .ActionFct = constValider
    .ActionAmj = DSys
    .ActionHms = time_Hms
    .ActionUsr = usrId
    .Statut = "F"
    .StatutPlus = "in"
End With

mearrGMemo_Nb = 0

For I = 1 To mearrGEch_Nb
    If mearrGEch(I).Statut = "à" And mearrGEch(I).StatutPlus = "C " Then
        meGECh = mearrGEch(I)
        V = GFlux_Scan(meGECh.FluxSéquence)
        If IsNull(V) Then
            With mearrGEch(I)
                .ActionFct = constCompta
                .ActionAmj = DSys
                .ActionHms = time_Hms
                .ActionUsr = usrId
            End With
            Call srvGSub_EC.GMemo_Gen(paramEffetCommerce, meGOpe, meGFlux, meGECh, warrGMemo_Nb, warrGMemo())
            For K = 1 To warrGMemo_Nb
                If mearrGMemo_Nb = mearrGMemo_NbMax Then mearrGMemo_NbMax = mearrGMemo_NbMax + 10: ReDim Preserve mearrGMemo(mearrGMemo_NbMax)

                mearrGMemo_Nb = mearrGMemo_Nb + 1
                mearrGMemo(mearrGMemo_Nb) = warrGMemo(K)
                mearrGMemo(mearrGMemo_Nb).Method = constAddNew
                mearrGMemo(mearrGMemo_Nb).MemoSéquence = 0                  ' séquence automatique AddNew
                mearrGMemo(mearrGMemo_Nb).MemoLien1 = paramEffetCommerce.ComptaLot
                mearrGMemo(mearrGMemo_Nb).Statut = "à"
                mearrGMemo(mearrGMemo_Nb).StatutPlus = "C "
            Next K
        End If
    End If
Next I

If mearrGMemo_Nb > 0 Then blnComptaAuto = True '''EffetCommerceAut.Comptabiliser

End Sub


Public Sub cmdControl_OpérationD_Devise()
Dim I1 As Integer, I2 As Integer
meCV1.DeviseIso = meGOpe.Devise1
V = CV_Attribut(meCV1): If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, "? erreur devise : " & meCV1.DeviseIso)
meGOpe.EngagementCompte = ""
meGOpe.EchéanceCompte = ""
meElpTable.Method = "Seek="
meElpTable.Id = paramEffetCommerce.TableId
meElpTable.K1 = "Devise"

meElpTable.K2 = meGOpe.Devise1
V = dbElpTable_ReadE(meElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(meElpTable.Memo) Then GoTo Table_Error
I1 = InStr(1, meElpTable.Memo, " ")
txtComEndos = Format$(Val(mId$(meElpTable.Memo, 1, I1)), "#0.000")
I2 = InStr(I1 + 1, meElpTable.Memo, " ")
txtComManipulation = Format$(Val(mId$(meElpTable.Memo, I1 + 1, I2 - I1 - 1)), "####0.00")

Exit Sub

Table_Error:
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Table", vbCritical, "frmEffetCommerce.cmdControl_OpéraionD_Devise"
Exit Sub

End Sub

Private Sub txtTiréDomiciliation_GotFocus()
txt_GotFocus txtTiréDomiciliation
End Sub

Private Sub txtTiréDomiciliation_LostFocus()
txt_LostFocus txtTiréDomiciliation
If blnControl Then cmdControl

End Sub


Private Sub txtTiréNom_GotFocus()
txt_GotFocus txtTiréNom

End Sub

Private Sub txtTiréNom_LostFocus()
txt_LostFocus txtTiréNom
If blnControl Then cmdControl

End Sub


Private Sub txtTiréRéférence_GotFocus()
txt_GotFocus txtTiréRéférence

End Sub

Private Sub txtTiréRéférence_LostFocus()
txt_LostFocus txtTiréRéférence
If blnControl Then cmdControl

End Sub





Public Sub fgSelect_MouseDown_Opération()
Dim xStatut As String

            mnuOpérationDisplay = EffetCommerceAut.Consulter
            mnuOpérationModifier = False
            mnuOpérationAnnuler = False
            mnuOpérationEffacer = False
            mnuOpérationValider = False
          
            xStatut = mGOpe.Statut & mGOpe.StatutPlus
            If xStatut = "à? " Then
                mnuOpérationModifier = EffetCommerceAut.Saisir
                mnuOpérationEffacer = EffetCommerceAut.Saisir
            End If
            If xStatut = "à  " Then
                mnuOpérationModifier = EffetCommerceAut.Saisir
                mnuOpérationEffacer = EffetCommerceAut.Saisir
            End If
            If xStatut = "àV " Then
              If Not EffetCommerceAut.Xspécial And Trim(mearrGEch(1).EchUsr) = Trim(usrId) Then
                    Call lstErr_Clear(lstErr, cmdContext, "! Vous ne pouvez pas valider vos opérations")
                Else
                    mnuOpérationValider = EffetCommerceAut.Valider
                End If
             End If
'$$$            If xStatut = "   " Then
'$$$                mnuTCAMJFin = EffetCommerceAut.Saisir
'$$$                mnuTCMainLevéePartielle = EffetCommerceAut.Saisir
'$$$                mnuTCMainLevée = EffetCommerceAut.Saisir
'$$$           End If
    
            Me.PopupMenu mnuOpération, vbPopupMenuLeftButton

End Sub
Public Sub fgSelect_MouseDown_Validation()

X = Chr$(9) & mGOpe.RéférenceExterne
mnuLotàComptabiliserValider.Caption = mLotàComptabiliserValider & X
mnuLotàComptabiliserAnnuler.Caption = mLotàComptabiliserAnnuler & X
mnuLotàComptabiliserPrint.Caption = mLotàComptabiliserPrint & X

mnuLotàComptabiliserValider = EffetCommerceAut.Valider
mnuLotàComptabiliserAnnuler = EffetCommerceAut.Valider
mnuLotàComptabiliserPrint = EffetCommerceAut.Consulter
         
    
Me.PopupMenu mnuLot, vbPopupMenuLeftButton

End Sub


Public Sub cmdPrint_Dossier()
Dim I As Integer

Me.Enabled = False
prtEffetCommerce_Dossier_Open


prtEffetCommerce_Dossier_GOpe paramEffetCommerce, meGOpe, mearrGEch(1), mearrGMemo(1)
prtEffetCommerce_Dossier_Form
For I = 2 To mearrGEch_Nb
     meGECh = mearrGEch(I)
    If Trim(meGECh.ActionFct) <> "" Then
        V = GMemo_Scan(meGECh.EchSéquence)
    Else
        V = GFlux_Scan(meGECh.FluxSéquence)
        If IsNull(V) Then
            Call srvGSub_EC.GMemo_Gen(paramEffetCommerce, meGOpe, meGFlux, meGECh, warrGMemo_Nb, warrGMemo())
        End If
    End If

    prtEffetCommerce_Dossier_GEch paramEffetCommerce, meGOpe, meGECh, warrGMemo_Nb, warrGMemo()
Next I

prtEffetCommerce_Close

Exit_Sub:

Me.Enabled = True
AppActivate Me.Caption


End Sub
