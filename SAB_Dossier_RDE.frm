VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSAB_Dossier_RDE 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_Dossier_RDE"
   ClientHeight    =   12165
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   16335
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SAB_Dossier_RDE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   12165
   ScaleWidth      =   16335
   Begin VB.ListBox lstPrinters 
      BackColor       =   &H00E0FFE0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      ItemData        =   "SAB_Dossier_RDE.frx":030A
      Left            =   11805
      List            =   "SAB_Dossier_RDE.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   53
      Top             =   90
      Width           =   3795
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00F0F0F0&
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
      Height          =   510
      Left            =   6255
      TabIndex        =   2
      Top             =   0
      Width           =   5490
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
      Height          =   11640
      Left            =   20
      TabIndex        =   3
      Top             =   540
      Width           =   16275
      Begin TabDlg.SSTab SSTab1 
         Height          =   11430
         Left            =   90
         TabIndex        =   4
         Top             =   90
         Width           =   16095
         _ExtentX        =   28390
         _ExtentY        =   20161
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Dossier"
         TabPicture(0)   =   "SAB_Dossier_RDE.frx":030E
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdPrint_Dossier"
         Tab(0).Control(1)=   "fraDossier"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Paramétrage"
         TabPicture(1)   =   "SAB_Dossier_RDE.frx":032A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sstabParam"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "."
         TabPicture(2)   =   "SAB_Dossier_RDE.frx":0346
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "SSTab2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.CommandButton cmdPrint_Dossier 
            BackColor       =   &H00C0FFC0&
            Caption         =   "enregistrer : dossier.pdf"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   -74850
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   90
            Visible         =   0   'False
            Width           =   2010
         End
         Begin VB.Frame fraDossier 
            Height          =   10890
            Left            =   -74970
            TabIndex        =   24
            Top             =   450
            Width           =   15945
            Begin SHDocVwCtl.WebBrowser WebBrowser1 
               Height          =   10695
               Left            =   120
               TabIndex        =   54
               Top             =   120
               Width           =   7755
               ExtentX         =   13679
               ExtentY         =   18865
               ViewMode        =   0
               Offline         =   0
               Silent          =   0
               RegisterAsBrowser=   0
               RegisterAsDropTarget=   1
               AutoArrange     =   0   'False
               NoClientEdge    =   0   'False
               AlignLeft       =   0   'False
               NoWebView       =   0   'False
               HideFileNames   =   0   'False
               SingleClick     =   0   'False
               SingleSelection =   0   'False
               NoFolders       =   0   'False
               Transparent     =   0   'False
               ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
               Location        =   "http:///"
            End
            Begin VB.ListBox lstCourrier 
               BackColor       =   &H00F0FFFF&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5160
               ItemData        =   "SAB_Dossier_RDE.frx":0362
               Left            =   8925
               List            =   "SAB_Dossier_RDE.frx":0369
               Style           =   1  'Checkbox
               TabIndex        =   31
               Top             =   120
               Width           =   6960
            End
            Begin VB.Frame fraCourrier 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Height          =   4215
               Left            =   7980
               TabIndex        =   28
               Top             =   1020
               Width           =   915
               Begin VB.OptionButton optCourrier_All 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "tous"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Left            =   180
                  TabIndex        =   43
                  Top             =   990
                  Width           =   700
               End
               Begin VB.OptionButton optCourrier_REG 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "REG"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Left            =   180
                  TabIndex        =   30
                  Top             =   570
                  Width           =   765
               End
               Begin VB.OptionButton optCourrier_OUV 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "OUV"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Left            =   180
                  TabIndex        =   29
                  Top             =   180
                  Value           =   -1  'True
                  Width           =   700
               End
            End
            Begin VB.Frame fraLangue 
               BackColor       =   &H00F0FFFF&
               BorderStyle     =   0  'None
               Height          =   840
               Left            =   7980
               TabIndex        =   25
               Top             =   165
               Width           =   915
               Begin VB.OptionButton optLangue_FR 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "FR"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   180
                  TabIndex        =   27
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   600
               End
               Begin VB.OptionButton optLangue_GB 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "GB"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   180
                  TabIndex        =   26
                  Top             =   480
                  Width           =   600
               End
            End
         End
         Begin TabDlg.SSTab sstabParam 
            Height          =   10290
            Left            =   -74985
            TabIndex        =   8
            Top             =   435
            Width           =   15615
            _ExtentX        =   27543
            _ExtentY        =   18150
            _Version        =   393216
            Tab             =   1
            TabHeight       =   520
            TabCaption(0)   =   "Paramètres SAB_Dossier_RDE"
            TabPicture(0)   =   "SAB_Dossier_RDE.frx":037A
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "CmDialog"
            Tab(0).Control(1)=   "btnImprimer"
            Tab(0).Control(2)=   "fraParam_BIATABK2"
            Tab(0).Control(3)=   "lstParam_BIATABK2"
            Tab(0).Control(4)=   "lstParam_BIATABK1"
            Tab(0).ControlCount=   5
            TabCaption(1)   =   "Gestion des modèles Word"
            TabPicture(1)   =   "SAB_Dossier_RDE.frx":0396
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "libParam_Modèles_Temp_Path"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "libParam_Modèles_REMDOC_Path"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "lstParam_Modèles_REMDOC"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "lstParam_Modèles_Temp"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "fraParam_Courrier"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "btnControle"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).ControlCount=   6
            TabCaption(2)   =   "Tableau récapitulatif"
            TabPicture(2)   =   "SAB_Dossier_RDE.frx":03B2
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "optParam_Recap_SAB"
            Tab(2).Control(1)=   "optParam_Recap_DDS"
            Tab(2).Control(2)=   "optParam_Recap_Z"
            Tab(2).Control(3)=   "fgParam_Recap"
            Tab(2).ControlCount=   4
            Begin VB.CommandButton btnControle 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Contrôle des variables non à jour"
               Height          =   555
               Left            =   6630
               Style           =   1  'Graphical
               TabIndex        =   58
               Top             =   750
               Width           =   2025
            End
            Begin MSComDlg.CommonDialog CmDialog 
               Left            =   -69390
               Top             =   300
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.CommandButton btnImprimer 
               Caption         =   "Imprimer"
               Height          =   345
               Left            =   -71160
               TabIndex        =   55
               Top             =   420
               Width           =   1305
            End
            Begin VB.OptionButton optParam_Recap_SAB 
               Caption         =   "champ #SAB / courriers"
               Height          =   255
               Left            =   -68205
               TabIndex        =   49
               Top             =   525
               Width           =   2310
            End
            Begin VB.OptionButton optParam_Recap_DDS 
               Caption         =   "caractéristiques des courriers"
               Height          =   255
               Left            =   -72390
               TabIndex        =   48
               Top             =   510
               Width           =   3030
            End
            Begin VB.OptionButton optParam_Recap_Z 
               Caption         =   "Aucun"
               Height          =   255
               Left            =   -74310
               TabIndex        =   47
               Top             =   450
               Value           =   -1  'True
               Width           =   1065
            End
            Begin VB.Frame fraParam_Courrier 
               Caption         =   "Indiquer les caractéristiques du courrier"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   8970
               Left            =   7380
               TabIndex        =   37
               Top             =   2040
               Visible         =   0   'False
               Width           =   6960
               Begin VB.TextBox txtParam_Courrier_Originaux 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080C0FF&
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2010
                  MaxLength       =   1
                  TabIndex        =   52
                  Text            =   "1"
                  Top             =   8460
                  Width           =   450
               End
               Begin VB.TextBox txtParam_Courrier_Seq 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FF80&
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   4125
                  MaxLength       =   4
                  TabIndex        =   41
                  Top             =   8505
                  Width           =   1035
               End
               Begin VB.CommandButton cmdParam_Courrier_Update 
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
                  Height          =   500
                  Left            =   5700
                  Style           =   1  'Graphical
                  TabIndex        =   40
                  Top             =   8300
                  Width           =   1110
               End
               Begin VB.CommandButton cmdParam_Courrier_Quit 
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
                  Left            =   135
                  Style           =   1  'Graphical
                  TabIndex        =   39
                  Top             =   8300
                  Width           =   1170
               End
               Begin MSFlexGridLib.MSFlexGrid fgParam_Courrier 
                  Height          =   7545
                  Left            =   210
                  TabIndex        =   38
                  Top             =   420
                  Visible         =   0   'False
                  Width           =   6585
                  _ExtentX        =   11615
                  _ExtentY        =   13309
                  _Version        =   393216
                  Cols            =   3
                  FixedCols       =   0
                  RowHeightMin    =   350
                  BackColor       =   15794175
                  ForeColor       =   16384
                  BackColorFixed  =   12648447
                  ForeColorFixed  =   16576
                  BackColorBkg    =   15794175
                  WordWrap        =   -1  'True
                  AllowUserResizing=   3
                  FormatString    =   "<N°      |<Description                                                                                  |<Indicateur"
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
               Begin VB.Label libParam_Courrier_Originaux 
                  BackColor       =   &H0080C0FF&
                  Caption         =   "Nb originaux (1-3)"
                  Height          =   255
                  Left            =   1545
                  TabIndex        =   51
                  Top             =   8070
                  Width           =   1440
               End
               Begin VB.Label libParam_Courrier_Seq 
                  BackColor       =   &H0080FF80&
                  Caption         =   "Ordre d'affichage des courriers sélectionnés"
                  Height          =   465
                  Left            =   3675
                  TabIndex        =   42
                  Top             =   7995
                  Width           =   1890
               End
            End
            Begin VB.Frame fraParam_BIATABK2 
               BackColor       =   &H0080C0FF&
               Height          =   4830
               Left            =   -66210
               TabIndex        =   15
               Top             =   2250
               Visible         =   0   'False
               Width           =   6585
               Begin VB.ComboBox cboParam_BIATABK2 
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   2535
                  Sorted          =   -1  'True
                  TabIndex        =   23
                  Text            =   "cboParam_BIATABK2"
                  Top             =   795
                  Width           =   2400
               End
               Begin VB.TextBox txtParam_BIATABTXT 
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1320
                  Left            =   690
                  MultiLine       =   -1  'True
                  TabIndex        =   21
                  Text            =   "SAB_Dossier_RDE.frx":03CE
                  Top             =   2220
                  Width           =   5355
               End
               Begin VB.TextBox txtParam_BIATABK2 
                  Height          =   350
                  Left            =   2565
                  MaxLength       =   12
                  TabIndex        =   20
                  Top             =   240
                  Width           =   2295
               End
               Begin VB.CommandButton cmdParam_Detail_Delete 
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
                  Height          =   480
                  Left            =   5055
                  Style           =   1  'Graphical
                  TabIndex        =   19
                  Top             =   3900
                  Width           =   900
               End
               Begin VB.CommandButton cmdParam_Detail_Add 
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
                  Height          =   480
                  Left            =   2160
                  Style           =   1  'Graphical
                  TabIndex        =   18
                  Top             =   3900
                  Width           =   900
               End
               Begin VB.CommandButton cmdParam_Detail_Update 
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
                  Height          =   480
                  Left            =   3645
                  Style           =   1  'Graphical
                  TabIndex        =   17
                  Top             =   3900
                  Width           =   900
               End
               Begin VB.CommandButton cmdParam_Detail_Quit 
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
                  Height          =   480
                  Left            =   405
                  Style           =   1  'Graphical
                  TabIndex        =   16
                  Top             =   3900
                  Width           =   990
               End
               Begin VB.Label lblParam_BIATABTXT 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Pour initialiser le contenu d'un champ variable de type ?...., mettre le texte entre guillemets : ""valeur par défaut"""
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   570
                  Left            =   750
                  TabIndex        =   50
                  Top             =   1425
                  Width           =   5205
               End
               Begin VB.Label lblParam_BIATABK2 
                  BackColor       =   &H0080C0FF&
                  Caption         =   "Code "
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
                  Left            =   540
                  TabIndex        =   22
                  Top             =   345
                  Width           =   1290
               End
            End
            Begin VB.ListBox lstParam_BIATABK2 
               BackColor       =   &H00E0FFFF&
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   9180
               Left            =   -71190
               TabIndex        =   14
               Top             =   870
               Width           =   11520
            End
            Begin VB.ListBox lstParam_BIATABK1 
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   9180
               Left            =   -74835
               Sorted          =   -1  'True
               TabIndex        =   13
               Top             =   855
               Width           =   3570
            End
            Begin VB.ListBox lstParam_Modèles_Temp 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7935
               Left            =   8160
               TabIndex        =   10
               Top             =   1710
               Width           =   7000
            End
            Begin VB.ListBox lstParam_Modèles_REMDOC 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7935
               Left            =   510
               TabIndex        =   9
               Top             =   1605
               Width           =   7000
            End
            Begin MSFlexGridLib.MSFlexGrid fgParam_Recap 
               Height          =   9030
               Left            =   -74790
               TabIndex        =   46
               Top             =   960
               Visible         =   0   'False
               Width           =   14865
               _ExtentX        =   26220
               _ExtentY        =   15928
               _Version        =   393216
               Cols            =   20
               FixedCols       =   0
               RowHeightMin    =   250
               BackColor       =   16777215
               ForeColor       =   16384
               BackColorFixed  =   8421504
               ForeColorFixed  =   -2147483633
               BackColorBkg    =   -2147483633
               AllowUserResizing=   3
               FormatString    =   "<Intitulé                             |   |   |   |   |   |   |   |   |   |   |   |   |   |   |   |   |   |   |"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label libParam_Modèles_REMDOC_Path 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Liste des modèles"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   495
               Left            =   690
               TabIndex        =   12
               Top             =   800
               Width           =   5760
            End
            Begin VB.Label libParam_Modèles_Temp_Path 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "C:\Temp\"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   495
               Left            =   8850
               TabIndex        =   11
               Top             =   800
               Width           =   5745
            End
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   10185
            Left            =   150
            TabIndex        =   5
            Top             =   405
            Width           =   15405
            _ExtentX        =   27173
            _ExtentY        =   17965
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "Tab 0"
            TabPicture(0)   =   "SAB_Dossier_RDE.frx":0403
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "txtFg"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "cboSelect_SQL"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "fraInfo_M"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Tab 1"
            TabPicture(1)   =   "SAB_Dossier_RDE.frx":041F
            Tab(1).ControlEnabled=   0   'False
            Tab(1).ControlCount=   0
            TabCaption(2)   =   "Tab 2"
            TabPicture(2)   =   "SAB_Dossier_RDE.frx":043B
            Tab(2).ControlEnabled=   0   'False
            Tab(2).ControlCount=   0
            Begin VB.Frame fraInfo_M 
               BackColor       =   &H00F0F0F0&
               Caption         =   "Informations complémentaires"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   10000
               Left            =   3300
               TabIndex        =   32
               Top             =   -420
               Width           =   11685
               Begin VB.TextBox txtUTI_DOC_M 
                  BackColor       =   &H0080C0FF&
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2385
                  MaxLength       =   2000
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   57
                  Top             =   1950
                  Visible         =   0   'False
                  Width           =   1155
               End
               Begin MSFlexGridLib.MSFlexGrid fgUTI_DOC 
                  Height          =   8445
                  Left            =   7770
                  TabIndex        =   56
                  Top             =   1500
                  Visible         =   0   'False
                  Width           =   7320
                  _ExtentX        =   12912
                  _ExtentY        =   14896
                  _Version        =   393216
                  Cols            =   5
                  FixedCols       =   0
                  RowHeightMin    =   350
                  BackColor       =   16448250
                  ForeColor       =   16384
                  BackColorFixed  =   33023
                  ForeColorFixed  =   -2147483633
                  BackColorBkg    =   16448250
                  AllowUserResizing=   3
                  FormatString    =   "<Document                                                                 |<1er Jeu            |<2ème Jeu       |>N°         "
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
               Begin VB.TextBox txtInfo_M 
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
                  Height          =   510
                  Left            =   1320
                  MaxLength       =   2000
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   36
                  Top             =   1290
                  Visible         =   0   'False
                  Width           =   4515
               End
               Begin VB.CommandButton cmdInfo_M_Ok 
                  BackColor       =   &H0080FF80&
                  Caption         =   "Continuer"
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
                  Left            =   6255
                  Style           =   1  'Graphical
                  TabIndex        =   35
                  Top             =   9210
                  Width           =   1560
               End
               Begin VB.CommandButton cmdInfo_M_Quit 
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
                  Left            =   1350
                  Style           =   1  'Graphical
                  TabIndex        =   33
                  Top             =   9135
                  Width           =   1605
               End
               Begin MSFlexGridLib.MSFlexGrid fgInfo_M 
                  Height          =   8520
                  Left            =   870
                  TabIndex        =   34
                  Top             =   885
                  Width           =   11445
                  _ExtentX        =   20188
                  _ExtentY        =   15028
                  _Version        =   393216
                  Cols            =   4
                  FixedCols       =   0
                  RowHeightMin    =   350
                  BackColor       =   15794175
                  ForeColor       =   16384
                  BackColorFixed  =   12648447
                  ForeColorFixed  =   16576
                  BackColorBkg    =   15794175
                  WordWrap        =   -1  'True
                  AllowUserResizing=   3
                  FormatString    =   $"SAB_Dossier_RDE.frx":0457
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
            End
            Begin VB.ComboBox cboSelect_SQL 
               Height          =   330
               Left            =   660
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   420
               Visible         =   0   'False
               Width           =   3225
            End
            Begin VB.TextBox txtFg 
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
               Height          =   3015
               Left            =   300
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   6
               Text            =   "SAB_Dossier_RDE.frx":0535
               Top             =   1335
               Visible         =   0   'False
               Width           =   5595
            End
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "cliquer ici pour fermer cette fenêtre (ESC = fermeture progressive)"
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   15600
      MaskColor       =   &H80000000&
      Picture         =   "SAB_Dossier_RDE.frx":053D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   -15
      Width           =   705
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   1245
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
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
   Begin VB.Menu mnuParam_Modèles_REMDOC 
      Caption         =   "mnuParam_Modèles_REMDOC"
      Visible         =   0   'False
      Begin VB.Menu mnuParam_Modèles_REMDOC_Des 
         Caption         =   "Caractéristiques du courrier"
      End
      Begin VB.Menu mnuParam_Modèles_REMDOC_Copier 
         Caption         =   "Copier vers le répertoire de travail"
      End
      Begin VB.Menu mnuParam_Modèles_REMDOC_Delete 
         Caption         =   "Supprimer du répertoire de production"
      End
      Begin VB.Menu mnuParam_Modèles_REMDOC_Rename 
         Caption         =   "Renommer un modèle"
      End
   End
   Begin VB.Menu mnuParam_Modèles_Temp 
      Caption         =   "mnuParam_Modèles_Temp"
      Visible         =   0   'False
      Begin VB.Menu mnuParam_Modèles_Temp_Copier 
         Caption         =   "Copier vers le répertoire de PRODUCTION"
      End
      Begin VB.Menu mnuParam_Modèles_Temp_Delete 
         Caption         =   "Supprimer du répertoire de travail"
      End
   End
   Begin VB.Menu mnuParam_Courrier 
      Caption         =   "mnuParam_Courrier"
      Visible         =   0   'False
      Begin VB.Menu mnuParam_Courrier_OK 
         Caption         =   "Oui"
      End
      Begin VB.Menu mnuParam_Courrier_NOK 
         Caption         =   "Non"
      End
      Begin VB.Menu mnuParam_Courrier_Z 
         Caption         =   "sans objet"
      End
   End
   Begin VB.Menu mnuPrint2 
      Caption         =   "mnuPrint2"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint2_Excel 
         Caption         =   "Excel"
      End
      Begin VB.Menu mnuPrint2_Mail 
         Caption         =   "Envoi Mail"
      End
   End
   Begin VB.Menu mnuExemplaires 
      Caption         =   "mnuExemplaires"
      Visible         =   0   'False
      Begin VB.Menu mnuExemplaires_1 
         Caption         =   "1 original"
      End
      Begin VB.Menu mnuExemplaires_2 
         Caption         =   "2 originaux"
      End
      Begin VB.Menu mnuExemplaires_3 
         Caption         =   "3 originaux"
      End
   End
End
Attribute VB_Name = "frmSAB_Dossier_RDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


'---------------------------------------------------------
Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim arrHab(19) As Boolean
Dim blnAuto As Boolean, blnError As Boolean
Dim cmdSelect_SQL_K As String
Dim wAmjMin As String, wAmjMax As String

Dim HeightOfLine As Long, LinesOfText As Long

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim fgParam_Courrier_FormatString As String, fgParam_Courrier_K As Integer
Dim fgParam_Courrier_RowDisplay As Integer, fgParam_Courrier_RowClick As Integer, fgParam_Courrier_ColClick As Integer
Dim fgParam_Courrier_ColorClick As Long, fgParam_Courrier_ColorDisplay As Long
Dim fgParam_Courrier_Sort1 As Integer, fgParam_Courrier_Sort2 As Integer
Dim fgParam_Courrier_SortAD As Integer, fgParam_Courrier_Sort1_Old As Integer
Dim fgParam_Courrier_arrIndex As Integer
Dim blnfgParam_Courrier_DisplayLine As Boolean

Dim fgParam_Recap_FormatString As String, fgParam_Recap_K As Integer
Dim fgParam_Recap_RowDisplay As Integer, fgParam_Recap_RowClick As Integer, fgParam_Recap_ColClick As Integer
Dim fgParam_Recap_ColorClick As Long, fgParam_Recap_ColorDisplay As Long
Dim fgParam_Recap_Sort1 As Integer, fgParam_Recap_Sort2 As Integer
Dim fgParam_Recap_SortAD As Integer, fgParam_Recap_Sort1_Old As Integer
Dim fgParam_Recap_arrIndex As Integer
Dim blnfgParam_Recap_DisplayLine As Boolean

Dim fgInfo_M_FormatString As String, fgInfo_M_K As Integer
Dim fgInfo_M_RowDisplay As Integer, fgInfo_M_RowClick As Integer, fgInfo_M_ColClick As Integer
Dim fgInfo_M_ColorClick As Long, fgInfo_M_ColorDisplay As Long
Dim fgInfo_M_Sort1 As Integer, fgInfo_M_Sort2 As Integer
Dim fgInfo_M_SortAD As Integer, fgInfo_M_Sort1_Old As Integer
Dim fgInfo_M_arrIndex As Integer
Dim blnfgInfo_M_DisplayLine As Boolean

Dim oldYBIATAB0 As typeYBIATAB0, newYBIATAB0 As typeYBIATAB0
Dim blnParam_Update As Boolean
Dim oldCourrier_Doc As typeYBIATAB0, newCourrier_Doc As typeYBIATAB0
Dim blnCourrier_Doc_Exist As Boolean
Dim oldCourrier_Des As typeYBIATAB0, newCourrier_Des As typeYBIATAB0
Dim blnCourrier_Des_Exist As Boolean

Dim arrCourrier_Doc() As String, arrCourrier_Doc_Nb As Integer
Dim arrCourrier_Originaux_Param_Nb() As Integer, arrCourrier_Originaux_Dossier_Nb() As Integer
Dim mnuExemplaires_K As Integer
Dim arrCourrier_Des() As String, arrCourrier_Id() As String
Dim mCourrier_Des As String, mCourrier_Des_Len As Integer

Dim mFct_Caller As String
Dim xZCDODOS0 As typeZCDODOS0, xYCDOTIE0 As typeZCDOTIE0, xZCDOUTI0 As typeZCDOUTI0
Dim rsSabX As New ADODB.Recordset

'DR 23/06/2014 REMDOC
Dim xZENCCAR0 As typeZENCCAR0
Dim xYENCTIE0 As typeZENCTIE0
Dim blnPrint_Courrier_Ok As Boolean
Dim mBEN_BQE As String 'Banque du bénéficiaire
Dim paramREMDOC_FAX As String
Dim paramREMDOC_TEL_NEGO As String
Dim xZENCREG0 As typeZENCREG0
Dim mMTD_NET As Currency
Dim mTOTAL_TTC As Currency
Dim mCOM1 As Currency
Dim mTVA_T As Currency
Dim mFRAIS1 As Currency
Dim mFRAIS2 As Currency

Dim mCOP_DOS As String
Dim mMON_DEV As String
Dim mCON As String, mEVE As String, mETA As String
Dim mOUV_VAL As String
Dim mX_OUI As String
Dim mX_NON As String
Dim mBQE_ZADRESS0 As typeZADRESS0
Dim mBQE_Concat As String
Dim mDON_ZADRESS0 As typeZADRESS0
Dim mDON_Concat As String
Dim mBEN_ZADRESS0 As typeZADRESS0, mBEN_CDOTIESRN As String
Dim mBEN_Concat As String
Dim mBED_ZADRESS0 As typeZADRESS0
Dim mBED_Concat As String
Dim blnBED_ZADRESSE0 As Boolean
Dim w_ZADRESSE0 As typeZADRESS0
Dim mREM_ZADRESS0 As typeZADRESS0
Dim mREM_Concat As String
Dim mSWISABSWID_103_Nb As Long

Dim mBQE_RBT_ZADRESS0 As typeZADRESS0, mBQE_RBT_Concat As String

Dim mMTD_T As String, mMTD_C As String, mMTD_D As String, mMTD_N As String
Dim mRatio_C As String, mRatio_N As String

Dim mBQE_RBT As String
Dim mSWISABSWID_707 As Long, mSWISABSWID_799 As Long
Dim mSWISABSWID_707_Nb As Long, mSWISABSWID_799_Nb As Long

Dim mTC2_X As String, mTC2_W As String, mTC2_C As String, mTC2_N As String
Dim mECNF As typeWCDOCOM0, mENOTIF As typeWCDOCOM0
Dim blnZCDOTCO0_CDE As Boolean
Dim mELVD As typeWCDOCOM0, mIDOCIR As typeWCDOCOM0, mEPDIF As typeWCDOCOM0, mEMODIF As typeWCDOCOM0
Dim mERFA As typeWCDOCOM0, mECSIL As typeWCDOCOM0
Dim WCDOCOM0_X As typeWCDOCOM0, mCOM_OUV As String

Dim mIOUV As typeWCDOCOM0, mILVD As typeWCDOCOM0, mIMODIF As typeWCDOCOM0, mIPDIF As typeWCDOCOM0
Dim mIRFA As typeWCDOCOM0, mIACD As typeWCDOCOM0, mIACCEP As typeWCDOCOM0
Dim blnZCDOTCO0_CDI As Boolean

Dim mANNEXES_NB As Long
Dim mDescription As String, mIrrégularités As String, mIrrégularités_Index As Integer
Dim blnZCDOUTI0_Select As Boolean

Dim mREG_DVA_CR As Long, mREG_DVA_DB As Long
Dim mAR_Accord As String, mAR_Courrier As String, mATTN As String

Dim appWord As Word.Application
Dim docWord As Word.Document
Dim mDOS_Path As String, mDOS_Id As String, mDOS_OPEN As Long, mDOS_OPEC As String
Dim mDOS_seq As Integer
Dim mDOS_File_pdf As String, mDOS_Modèle As String, mWord_PDF_Path As String
Dim arrDOC() As String, arrDoc_Nb As Integer
Dim arrDOC_FileName() As String, arrDOC_REF() As String, mCourrier_Id As String
Dim arrDoc_Originaux_Nb() As Integer
Dim mDoc_Page_Nb As Integer

Dim arrFields_SAB_Name() As String, arrFields_SAB_Value() As String, arrFields_SAB_Nb As Integer
Dim blnFields_SAB_Name() As Boolean
Dim arrFields_BIA_Name() As String, arrFields_BIA_Value() As String, arrFields_BIA_Nb As Integer
Dim blnFields_BIA_Name() As Boolean, arrFields_BIA_Lib() As String, arrFields_BIA_Index As Integer

Dim mUTI_DOC_Index As Integer, blnUTI_DOC_Ok As Boolean, mUTI_DOC_Col  As Integer
Dim blnUTI_DOC_Loaded As Boolean
Dim arrUTI_DOC_Tbl_Nb As Integer, arrUTI_COM_CR_Tbl_Nb As Integer, arrUTI_COM_DB_Tbl_Nb As Integer
Dim arrUTI_COM_Escompte_Tbl_Nb As Integer, arrUTI_COM_Escompte_Tbl(100) As Integer
Dim arrUTI_DOC_Tbl(100) As Integer, arrUTI_COM_CR_Tbl(100) As Integer, arrUTI_COM_DB_Tbl(100) As Integer
Dim arrUTI_BLOCAGE_Tbl(100) As Integer, arrUTI_BLOCAGE_Tbl_Nb As Integer
Dim curUTI_BLOCAGE As Currency

'___________________________________________________________________________________________
Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim mXls1_Row As Long, mXls1_Cols As Long, mXls1_File As Integer
Dim mXls1_Col_1 As Long, mXls1_Col_2 As Long

Dim xrText As typerText
Dim cnSIDE_DB As New ADODB.Connection, rsSIDE_DB As New ADODB.Recordset, blnSIDE_DB_Open As Boolean

Dim mWord_ActivePrinter As String

Dim hwndWord As Long
Dim mClipBoard

Private Sub cmdPrint_Courrier_Word_UTI_DOC(lUTI_DOC_Tbl As Integer)
Dim K As Integer, wRows_Count As Integer, wTab_Row As Integer
Dim wFont_Name As String, wFont_Size As Integer
On Error GoTo Error_Handler

wFont_Name = appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(1, 1).Range.Font.Name
wFont_Size = appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(1, 1).Range.Font.Size

wRows_Count = appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Rows.Count: wTab_Row = 1
    For K = 1 To fgUTI_DOC.Rows - 1
        fgUTI_DOC.Row = K
        fgUTI_DOC.Col = 0
        If fgUTI_DOC.CellBackColor = mColor_G1 Then
            wTab_Row = wTab_Row + 1
            If wTab_Row > wRows_Count Then
                wRows_Count = wRows_Count + 1
                appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Rows.Add
                'suppression de la dernière ligne vide en bas de la page
                Call Word_supprime_derniere_ligne_vide
            End If
            appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(wTab_Row, 1).Range.Text = Trim(fgUTI_DOC.Text)
            appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(wTab_Row, 1).Range.Font.Name = wFont_Name
            appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(wTab_Row, 1).Range.Font.Size = wFont_Size
            
            fgUTI_DOC.Col = 1: appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(wTab_Row, 2).Range.Text = Trim(fgUTI_DOC.Text)
            appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(wTab_Row, 2).Range.Font.Name = wFont_Name
            appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(wTab_Row, 2).Range.Font.Size = wFont_Size
            
            fgUTI_DOC.Col = 2: appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(wTab_Row, 3).Range.Text = Trim(fgUTI_DOC.Text)
            appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(wTab_Row, 3).Range.Font.Name = wFont_Name
            appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(wTab_Row, 3).Range.Font.Size = wFont_Size

            If wTab_Row Mod 2 = 0 Then
                appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(wTab_Row, 1).Shading.BackgroundPatternColor = wdColorGray05
                appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(wTab_Row, 2).Shading.BackgroundPatternColor = wdColorGray05
                appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(wTab_Row, 3).Shading.BackgroundPatternColor = wdColorGray05
            Else
                appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(wTab_Row, 1).Shading.BackgroundPatternColor = wdColorWhite
                appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(wTab_Row, 2).Shading.BackgroundPatternColor = wdColorWhite
                appWord.ActiveDocument.Tables(lUTI_DOC_Tbl).Cell(wTab_Row, 3).Shading.BackgroundPatternColor = wdColorWhite
            End If

        End If
        
    Next K

GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbInformation, "cmdPrint_Courrier_Word_UTI_DOC"
Exit_sub:
End Sub
Private Sub Word_supprime_derniere_ligne_vide()
Dim ii As Long

    With appWord.ActiveDocument
        Dim mps As Paragraphs
        Dim mp As Paragraph
        Set mps = .Paragraphs
        For ii = mps.Count - 1 To 1 Step -1
            If Trim(mps(ii).Range.Text) = "" Or Trim(mps(ii).Range.Text) = Chr(13) Then
                mps(ii).Range.Delete
                Exit For
            End If
        Next ii
    End With
    
End Sub

Public Sub fgUTI_DOC_Load()
Dim X As String, xSql As String, wN As Integer
On Error GoTo Error_Handler
currentAction = "fgUTI_DOC_Load"
fgUTI_DOC.Rows = 1
wN = 100
    xSql = "select * from " & paramIBM_Library_SAB & ".ZCDOTAB0" _
       & " where CDOTABETA = 1" _
       & " and CDOTABNUM = 19 and CDOTABARG like '004%'" _
       & " order by CDOTABDON"
    Set rsSabX = cnsab.Execute(xSql)
    Do While Not rsSabX.EOF
        fgUTI_DOC.Rows = fgUTI_DOC.Rows + 1
        fgUTI_DOC.Row = fgUTI_DOC.Rows - 1
        fgUTI_DOC.Col = 0
        fgUTI_DOC.Text = Trim(rsSabX("CDOTABDON"))
        fgUTI_DOC.Col = 4
        fgUTI_DOC.Text = Trim(rsSabX("CDOTABARG"))
        fgUTI_DOC.Col = 3
        
        If optLangue_FR Then
            Select Case Trim(Mid$(rsSabX("CDOTABARG"), 1, 15))
                Case "004FACT": fgUTI_DOC.Text = "0010"
                Case "004BL": fgUTI_DOC.Text = "0020"
                Case "004LTA": fgUTI_DOC.Text = "0030"
                Case "004NOTPOI": fgUTI_DOC.Text = "0040"
                Case "004LISCOL": fgUTI_DOC.Text = "0050"
                Case Else
                    wN = wN + 10
                    fgUTI_DOC.Text = Format(wN, "0000")
            End Select
        Else
            Select Case Trim(Mid$(rsSabX("CDOTABARG"), 1, 15))
                Case "004GFACT": fgUTI_DOC.Text = "0010"
                Case "004GBL": fgUTI_DOC.Text = "0020"
                Case "004GLTA": fgUTI_DOC.Text = "0030"
                Case "004GPOIDS": fgUTI_DOC.Text = "0040"
                Case "004GCOLIS": fgUTI_DOC.Text = "0050"
                Case Else
                    wN = wN + 10
                    fgUTI_DOC.Text = Format(wN, "0000")
            End Select
        End If
        
        
        rsSabX.MoveNext
    Loop
fgUTI_DOC_Sort
blnUTI_DOC_Loaded = True
Exit Sub
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub fgUTI_DOC_Sort()

    If fgUTI_DOC.Rows > 1 Then
        fgUTI_DOC.Row = 1
        fgUTI_DOC.RowSel = fgUTI_DOC.Rows - 1
        
        fgUTI_DOC.Col = 3
        fgUTI_DOC.ColSel = 3
        fgUTI_DOC.Sort = 5
    End If

End Sub

Public Sub cmdPrint_Courrier_Info_M()
Dim X As String, blnInfo_M_Ok As Boolean, K As Integer
On Error GoTo Error_Handler

currentAction = "cmdPrint_Courrier_Info_M"
fgInfo_M.Rows = 1
fgInfo_M.FormatString = fgInfo_M_FormatString
fgInfo_M.Row = 0
blnInfo_M_Ok = True
For K = 1 To arrFields_BIA_Nb
    X = Trim(arrFields_BIA_Name(K))
    With appWord.Selection.Find
        .Wrap = wdFindContinue
        .Text = X
        .Execute
    End With
    If appWord.Selection.Find.Found Then
         fgInfo_M.Rows = fgInfo_M.Rows + 1
         fgInfo_M.Row = fgInfo_M.Rows - 1
         fgInfo_M.RowHeight(fgInfo_M.Row) = 700
         fgInfo_M.Col = 0: fgInfo_M.Text = X: fgInfo_M.CellForeColor = vbMagenta
         fgInfo_M.Col = 1: fgInfo_M.Text = arrFields_BIA_Lib(K)
         fgInfo_M.Col = 3: fgInfo_M.Text = K
         fgInfo_M.Col = 2: fgInfo_M.CellBackColor = mColor_W0
         If arrFields_BIA_Value(K) <> "" Then
            fgInfo_M.Col = 2
            fgInfo_M.CellBackColor = mColor_G0
            fgInfo_M.Text = arrFields_BIA_Value(K)
         Else
            Select Case X
                Case "?ATTN":
                        fgInfo_M.Col = 2
                        fgInfo_M.CellBackColor = mColor_G0
                        fgInfo_M.Text = mATTN
                Case "?TELECOPIE":
                        fgInfo_M.Col = 2
                        fgInfo_M.CellBackColor = mColor_G0
                        fgInfo_M.Text = arrFields_BIA_Lib(K)
                Case "?UTI_DOC"
                      mUTI_DOC_Index = K
                      fgInfo_M.Col = 2
                      
                      If blnUTI_DOC_Ok Then
                          fgInfo_M.CellBackColor = mColor_G0
                      Else
                          blnInfo_M_Ok = False
                      End If
                Case Else
                        blnInfo_M_Ok = False
            End Select
        End If
    End If
Next K
Clipboard.Clear
Clipboard.SetText mClipBoard
fgInfo_M.Visible = True
fraInfo_M.Visible = True
If blnInfo_M_Ok Then cmdInfo_M_Ok.Visible = True
Exit Sub

Error_Handler:

Call MsgBox(Error, vbCritical, currentAction)

End Sub

Private Sub cmdPrint_Courrier_Word_Quit()

On Error GoTo Error_Handler
blnPrint_Courrier_Ok = False

' Close the document and Word.
Do Until appWord.Documents.Count = 0
         'Close no save
        appWord.Documents(1).Close SaveChanges:=wdDoNotSaveChanges
Loop

appWord.Quit False
cmdPrint.Visible = True

GoTo Exit_sub

Error_Handler:
    MsgBox Error
    appWord.Quit False
Exit_sub:
   Set docWord = Nothing
   Set appWord = Nothing
   DestroyWindow hwndWord

Clipboard.Clear
Clipboard.SetText mClipBoard

    Call lstErr_AddItem(lstErr, cmdContext, "Fermeture Word " & hwndWord): DoEvents
End Sub

Private Sub cmdPrint_Courrier_Word()
Dim X As String, K As Integer, K1 As Integer, xADR As String
Dim blnWord_Validation As Boolean, blnWord_Update As Boolean, blnPDF_Display As Boolean
Dim mPrinter_Word_Name As String, wXXX_OK As Integer, wXXX_NOK As Integer
On Error GoTo Error_Handler

currentAction = "cmdPrint_Courrier_Word"
blnWord_Update = False
blnWord_Validation = False
blnPDF_Display = False
blnPDF_Display = True
blnWord_Validation = True
blnWord_Validation = True: blnWord_Update = True
If fgInfo_M.Rows > 1 Then Call cmdPrint_Courrier_Info_M_Replace
Call lstErr_AddItem(lstErr, cmdContext, "Recherche variables #SAB ...."): DoEvents
ProgressBar1.value = ProgressBar1.value + 1
'_____________________________________________________________________________________
For K = 1 To UBound(arrFields_SAB_Value)
    arrFields_SAB_Value(K) = ""
    blnFields_SAB_Name(K) = False
Next K
mCourrier_Id = ""
For K = 1 To arrDoc_Nb
    For K1 = 1 To arrCourrier_Doc_Nb
        If arrDOC_FileName(K) = arrCourrier_Doc(K1) Then
            mCourrier_Id = mCourrier_Id & " , '" & arrCourrier_Id(K1) & "'"
            If InStr(arrDOC_FileName(K), " XXX ") > 0 Then
                wXXX_OK = wXXX_OK + 1
            Else
                wXXX_NOK = wXXX_NOK + 1
            End If
            Exit For
        End If
    Next K1
Next K
If wXXX_OK > 0 And wXXX_NOK = 0 Then blnWord_Validation = True: blnWord_Update = True
If mCourrier_Id = "" Then
    mCourrier_Id = "'???'" '"
Else
    mCourrier_Id = Replace(mCourrier_Id, " , ", "", 1, 1)
End If
X = "select distinct BIATABK2 from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'REMDOC_#SAB'" _
     & " and BIATABK1 in  (" & mCourrier_Id & ") order By BIATABK2"
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
        For K = 1 To arrFields_SAB_Nb
            If Trim(rsSab("BIATABK2")) = arrFields_SAB_Name(K) Then
                blnFields_SAB_Name(K) = True
                Exit For
            End If
        Next K
    Select Case arrFields_SAB_Name(K)
        Case "#COP_DOS":   arrFields_SAB_Value(K) = Trim(mCOP_DOS)
        Case "#FAX":        arrFields_SAB_Value(K) = Trim(paramREMDOC_FAX)
        Case "#JJ-MM-AAAA": arrFields_SAB_Value(K) = Trim(dateImp10_S(DSys))
        Case "#JJMOISAAAA": arrFields_SAB_Value(K) = convertDate_FRtoGB(optLangue_GB.value, Format(dateAAAAMMJJTOJJ_MM_AAAA(DSys), "dd mmmm yyyy"))
        Case "#TEL_NEGO": arrFields_SAB_Value(K) = Trim(paramREMDOC_TEL_NEGO)
        Case "#DEVISE", "#DEV_COM1", "#DEV_FRAIS1", "#DEV_FRAIS2": arrFields_SAB_Value(K) = Trim(xZENCCAR0.ENCCARDEV)
        Case "#MONTANT": arrFields_SAB_Value(K) = Format(xZENCCAR0.ENCCARMON, "### ### ### ##0.00")
        Case "#VALIDITE": arrFields_SAB_Value(K) = Trim(dateImp10_S(xZENCCAR0.ENCCARDAC + 19000000))
        Case "#BEN_BQE": arrFields_SAB_Value(K) = Trim(mBEN_BQE)
        Case "#BEN_CONCAT": arrFields_SAB_Value(K) = Trim(mBEN_Concat)
        Case "#BEN_RS1":    arrFields_SAB_Value(K) = Trim(mBEN_ZADRESS0.ADRESSRA1)
        Case "#BEN_RS2": arrFields_SAB_Value(K) = Trim(mBEN_ZADRESS0.ADRESSRA2)
        Case "#BEN_ADR1": arrFields_SAB_Value(K) = Trim(mBEN_ZADRESS0.ADRESSAD1)
        Case "#BEN_ADR2": arrFields_SAB_Value(K) = Trim(mBEN_ZADRESS0.ADRESSAD2)
        Case "#BEN_ADR3": arrFields_SAB_Value(K) = Trim(mBEN_ZADRESS0.ADRESSAD3)
        Case "#BEN_CP_VILL": arrFields_SAB_Value(K) = Trim(Trim(mBEN_ZADRESS0.ADRESSCOP) & " " & Trim(mBEN_ZADRESS0.ADRESSVIL))
        Case "#BEN_VILL": arrFields_SAB_Value(K) = Trim(mBEN_ZADRESS0.ADRESSVIL)
        Case "#BEN_PAYS": arrFields_SAB_Value(K) = Trim(mBEN_ZADRESS0.ADRESSPAY)
        Case "#DON_CONCAT": arrFields_SAB_Value(K) = Trim(mDON_Concat)
        Case "#DON_RS1": arrFields_SAB_Value(K) = Trim(mDON_ZADRESS0.ADRESSRA1)
        Case "#DON_RS2": arrFields_SAB_Value(K) = Trim(mDON_ZADRESS0.ADRESSRA2)
        Case "#DON_ADR1": arrFields_SAB_Value(K) = Trim(mDON_ZADRESS0.ADRESSAD1)
        Case "#DON_ADR2": arrFields_SAB_Value(K) = Trim(mDON_ZADRESS0.ADRESSAD2)
        Case "#DON_ADR3": arrFields_SAB_Value(K) = Trim(mDON_ZADRESS0.ADRESSAD3)
        Case "#DON_CP_VILL": arrFields_SAB_Value(K) = Trim(Trim(mDON_ZADRESS0.ADRESSCOP) & " " & Trim(mDON_ZADRESS0.ADRESSVIL))
        Case "#DON_VILL": arrFields_SAB_Value(K) = Trim(mDON_ZADRESS0.ADRESSVIL)
        Case "#DON_PAYS": arrFields_SAB_Value(K) = Trim(mBEN_ZADRESS0.ADRESSPAY)
        Case "#BQE_REF", "#REFEXT": arrFields_SAB_Value(K) = Trim(xZENCCAR0.ENCCARREX)
        Case "#BQE_RSX": arrFields_SAB_Value(K) = Trim(Trim(mBQE_ZADRESS0.ADRESSRA1) & Trim(mBQE_ZADRESS0.ADRESSRA2))
        Case "#BQE_ZIP": arrFields_SAB_Value(K) = LTrim(Trim(mBQE_ZADRESS0.ADRESSCOP) & " " & Trim(mBQE_ZADRESS0.ADRESSVIL) & " " & Trim(mBQE_ZADRESS0.ADRESSPAY))
        Case "#BQE_CONCAT": arrFields_SAB_Value(K) = Trim(mBQE_Concat)
        Case "#BQE_RS1": arrFields_SAB_Value(K) = Trim(mBQE_ZADRESS0.ADRESSRA1)
        Case "#BQE_RS2": arrFields_SAB_Value(K) = Trim(mBQE_ZADRESS0.ADRESSRA2)
        Case "#BQE_ADR1": arrFields_SAB_Value(K) = Trim(mBQE_ZADRESS0.ADRESSAD1)
        Case "#BQE_ADR2": arrFields_SAB_Value(K) = Trim(mBQE_ZADRESS0.ADRESSAD2)
        Case "#BQE_ADR3": arrFields_SAB_Value(K) = Trim(mBQE_ZADRESS0.ADRESSAD3)
        Case "#BQE_CP_VILL": arrFields_SAB_Value(K) = Trim(Trim(mBQE_ZADRESS0.ADRESSCOP) & " " & Trim(mBQE_ZADRESS0.ADRESSVIL))
        Case "#BQE_VILL": arrFields_SAB_Value(K) = Trim(mBQE_ZADRESS0.ADRESSVIL)
        Case "#BQE_PAYS": arrFields_SAB_Value(K) = Trim(mBQE_ZADRESS0.ADRESSPAY)
        Case "#REFEXT": arrFields_SAB_Value(K) = Trim(xZENCCAR0.ENCCARREX)
        Case "#REFCLI": arrFields_SAB_Value(K) = Trim(xZENCCAR0.ENCCARRCL)
        Case "#MTD_NET": arrFields_SAB_Value(K) = Trim(Format(mMTD_NET, "### ### ##0.00"))
        Case "#M_COM1": arrFields_SAB_Value(K) = Trim(Format(mCOM1, "### ### ##0.00"))
        Case "#M_FRAIS1": arrFields_SAB_Value(K) = Trim(Format(mFRAIS1, "### ### ##0.00"))
        Case "#M_FRAIS2": arrFields_SAB_Value(K) = Trim(Format(mFRAIS2, "### ### ##0.00"))
        Case "#TVA_T": arrFields_SAB_Value(K) = Trim(Format(mTVA_T, "### ### ##0.00"))
        Case "#TOTAL_TTC": arrFields_SAB_Value(K) = Trim(Format(mTOTAL_TTC, "### ### ### ##0.00"))
        Case "#DATE_ECHEAN": arrFields_SAB_Value(K) = Trim(dateImp10_S(xZENCCAR0.ENCCAREC1 + 19000000))
        Case "#DATE_VALEUR": arrFields_SAB_Value(K) = retourne_date_valeur(CLng(xZENCCAR0.ENCCARDOS))
    End Select
    rsSab.MoveNext
Loop
appWord.Selection.WholeStory
For K = 1 To arrFields_SAB_Nb
        With appWord.Selection.Find
            .Text = arrFields_SAB_Name(K)
            .Replacement.Text = Trim(arrFields_SAB_Value(K))
            .Execute Replace:=wdReplaceAll
        End With
Next K
ProgressBar1.value = ProgressBar1.value + 1
appWord.Selection.HomeKey Unit:=wdStory
appWord.ActivePrinter = mWord_ActivePrinter
ProgressBar1.value = ProgressBar1.value + 1
If blnWord_Validation Then
    If Not blnWord_Update Then appWord.ActiveDocument.Protect Type:=wdAllowOnlyReading, NoReset:=True
    appWord.Visible = True
    appWord.Windows.Application.Activate
    appWord.Windows.Application.WindowState = wdWindowStateMaximize
    hwndWord = FindWindow(vbNullString, "Microsoft Word")
    If hwndWord <> 0 Then
        Dim hwnd As Long
        hwnd = SetForegroundWindow(hwndWord)
    Else
       MsgBox "Impossible de trouver la fenêtre Word!", vbExclamation
    End If
    Sleep 2000
    X = MsgBox("Voulez_vous enregistrer ce document dans l'historique du courrier du dossier ?", vbYesNo, "SAB_Dossier_RDE")
    If Not blnWord_Update Then appWord.ActiveDocument.Unprotect
Else
    X = vbYes
End If
'$JPL 2013-01-08 ______________________________________________
If appWord.ActivePrinter <> mWord_ActivePrinter Then
    mWord_ActivePrinter = appWord.ActivePrinter
    lstPrinters_Load
End If
'$JPL 2013-01-08 ______________________________________________
If X = vbYes Then
    Call lstErr_AddItem(lstErr, cmdContext, "Impression Client...."): DoEvents
    appWord.PrintOut
    Call lstErr_AddItem(lstErr, cmdContext, "Impression Dossier...."): DoEvents
    Call docWord_Filigrane(appWord, mCOP_DOS, WdColor.wdColorSeaGreen) ' .wdColorBlueGray)
    appWord.PrintOut , , "3", , "1", CStr(mDoc_Page_Nb)
    If Dir(mWord_PDF_Path) <> "" Then
        Call docWord_Filigrane(appWord, "Duplicata", WdColor.wdColorSeaGreen)
        X = xZENCCAR0.ENCCARCOP & "_" & Format(xZENCCAR0.ENCCARDOS, "000000")
        mDOS_Path = paramRDE_Dossier_Path & X
        If Not msFileSystem.FolderExists(mDOS_Path) Then MkDir mDOS_Path
        mDOS_seq = mDOS_seq + 1
        mDOS_File_pdf = mDOS_Path & "\" & X & "_" & DSYS_Time & mDOS_seq & "_" & mDOS_Modèle & ".pdf"
        Call lstErr_AddItem(lstErr, cmdContext, "Enregistrement .pdf ...."): DoEvents
            Call appWord.ActiveDocument.ExportAsFixedFormat(mDOS_File_pdf, wdExportFormatPDF, blnPDF_Display, wdExportOptimizeForPrint)
            frmSAB_Dossier.fgCourrier_Display
            frmSAB_Dossier_RDE.Hide
   Else
       MsgBox "PDF add-in Not Installed"
   '------------------------------------------------------------------------
   End If
End If

GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbCritical, currentAction
Exit_sub:
    ProgressBar1.Visible = False
    cmdPrint_Courrier_Word_Quit
End Sub










Private Function convertDate_FRtoGB(lngGB As Boolean, z As String) As String

    convertDate_FRtoGB = LCase(z)
    If lngGB Then
        convertDate_FRtoGB = Replace(convertDate_FRtoGB, "janvier", "january")
        convertDate_FRtoGB = Replace(convertDate_FRtoGB, "février", "february")
        convertDate_FRtoGB = Replace(convertDate_FRtoGB, "mars", "march")
        convertDate_FRtoGB = Replace(convertDate_FRtoGB, "avril", "april")
        convertDate_FRtoGB = Replace(convertDate_FRtoGB, "mai", "may")
        convertDate_FRtoGB = Replace(convertDate_FRtoGB, "juin", "june")
        convertDate_FRtoGB = Replace(convertDate_FRtoGB, "juillet", "july")
        convertDate_FRtoGB = Replace(convertDate_FRtoGB, "août", "august")
        convertDate_FRtoGB = Replace(convertDate_FRtoGB, "septembre", "september")
        convertDate_FRtoGB = Replace(convertDate_FRtoGB, "octobre", "october")
        convertDate_FRtoGB = Replace(convertDate_FRtoGB, "novembre", "november")
        convertDate_FRtoGB = Replace(convertDate_FRtoGB, "décembre", "december")
    End If
    
End Function


Public Sub Form_Init(lFct As String, lCDODOSCOP As String, lCDODOSDOS As Long)
Dim X As String, K As Integer

On Error Resume Next
Call BIA_VB_HAB("SAB_DOS_CDO", arrHab(), cboSelect_SQL)
If Not arrHab(1) Then Unload Me: Exit Sub

sstabParam.Tab = 0
SSTab1.Tab = 0

WebBrowser1.Visible = False
lstCourrier.Clear
lstCourrier.Visible = False
sstabParam.Visible = arrHab(18)
fraInfo_M.Visible = False
'_______________________________________________________________________________________________________
X = "select count(*)  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'REMDOC' and BIATABK1 = '#SAB'"
Set rsSabX = cnsab.Execute(X)
If Not rsSabX.EOF Then
    ReDim arrFields_SAB_Name(rsSabX(0) + 10), arrFields_SAB_Value(rsSabX(0) + 10), blnFields_SAB_Name(rsSabX(0) + 10)
    arrFields_SAB_Nb = 0
    X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
          & " where BIATABID = 'REMDOC' and BIATABK1 = '#SAB' order by BIATABK2"
    Set rsSabX = cnsab.Execute(X)
    Do Until rsSabX.EOF
        arrFields_SAB_Nb = arrFields_SAB_Nb + 1
        arrFields_SAB_Name(arrFields_SAB_Nb) = Trim(rsSabX("BIATABK2"))
        rsSabX.MoveNext
    Loop
End If
'_______________________________________________________________________________________________________
X = "select count(*)  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'REMDOC' and BIATABK1 = '?BIA'"
Set rsSabX = cnsab.Execute(X)
If Not rsSabX.EOF Then
    ReDim arrFields_BIA_Name(rsSabX(0) + 10), arrFields_BIA_Value(rsSabX(0) + 10), blnFields_BIA_Name(rsSabX(0) + 10), arrFields_BIA_Lib(rsSabX(0) + 10)
    arrFields_BIA_Nb = 0
    X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
          & " where BIATABID = 'REMDOC' and BIATABK1 = '?BIA' order by BIATABK2"
    Set rsSabX = cnsab.Execute(X)
    Do Until rsSabX.EOF
        arrFields_BIA_Nb = arrFields_BIA_Nb + 1
        arrFields_BIA_Name(arrFields_BIA_Nb) = Trim(rsSabX("BIATABK2"))
        arrFields_BIA_Lib(arrFields_BIA_Nb) = Trim(rsSabX("BIATABTXT"))
        If arrFields_BIA_Name(arrFields_BIA_Nb) = "?UTI_IRR" Then mIrrégularités_Index = arrFields_BIA_Nb
        rsSabX.MoveNext
    Loop
End If
'_______________________________________________________________________________________________________
If arrHab(18) Then lstParam_BIATABK1_Init
If arrHab(2) Then
    If paramREMDOC_TEL_NEGO = "" Then paramREMDOC_Init
End If
cmdPrint.Visible = arrHab(2)
If Not frmSAB_Dossier_RDE.Visible Then frmSAB_Dossier_RDE.Visible = True
Me.Show
mFct_Caller = lFct
    X = "select *  from " & paramIBM_Library_SAB & ".ZENCCAR0 " _
          & " where ENCCARETA = " & currentSAB_ETA & " and ENCCARAGE = " & currentSAB_AGE _
          & " and ENCCARSER = '00' and ENCCARSSE = '00'" _
          & " and ENCCARCOP = '" & lCDODOSCOP & "' and ENCCARDOS = " & lCDODOSDOS
     Set rsSabX = cnsab.Execute(X)
    If rsSabX.EOF Then
        Call MsgBox("Dossier inconnu : " & lCDODOSCOP & " " & lCDODOSDOS, vbCritical, "SAB_Dossier_RDE")
        Exit Sub
    Else
        Call lstErr_Clear(lstErr, cmdContext, "Lecture du dossier : " & lCDODOSCOP & " " & lCDODOSDOS): DoEvents
        lstErr.Height = 510
        Call rsZENCCAR0_GetBuffer(rsSabX, xZENCCAR0)
        Call lstErr_AddItem(lstErr, cmdContext, "Affichage du dossier "): DoEvents
        Call fraDossier_Display
        optLangue_FR.value = True
        optCourrier_All.value = True
        If arrHab(2) Then lstCourrier_Load
    End If
Call lstPrinters_Load
X = Trim(frmSAB_Dossier_RDE.Caption)
AppActivate X
End Sub


Public Sub fraDossier_Display()
Dim X As String, curX As Currency, iRatio As Integer, K As Integer, K2 As Integer
Dim newdocumentRDE As String

optCourrier_OUV = False
mCOP_DOS = xZENCCAR0.ENCCARCOP & " " & Format(xZENCCAR0.ENCCARDOS, "### 000")
mMON_DEV = Format(xZENCCAR0.ENCCARMON, "### ### ### ###.00") & " " & xZENCCAR0.ENCCARDEV
mMTD_NET = 0
mBEN_BQE = ""
Select Case xZENCCAR0.ENCCARCET
    Case "01": mETA = "Saisi"
    Case "02": mETA = "Saisi - Validé"
    Case "03": mETA = "Saisie - Comptabilisé"
    Case "11": mETA = "Modifié"
    Case "12": mETA = "Modifié - Validé"
    Case "13": mETA = "Modifié - Comptabilisé"
    Case "91": mETA = "Clôturé - Saisi"
    Case "92": mETA = "Clôturé - Validé"
    Case "93": mETA = "Clôturé - Comptabilisé"
    Case Else: mETA = xZENCCAR0.ENCCARCET
End Select

'Lecture BQE agence - Banque présentatrice
X = Space(64)
'DR 01/06/2017
'Call rsZENCTIE_Adresse(Left(xZENCCAR0.ENCCARINT, 1), Mid(xZENCCAR0.ENCCARINT, 2), X, xYENCTIE0, mBQE_ZADRESS0, mBQE_Concat, "CD")
Call rsZENCTIE_Adresse(Left(xZENCCAR0.ENCCARINT, 1), Mid(xZENCCAR0.ENCCARINT, 2), X, xYENCTIE0, mBQE_ZADRESS0, mBQE_Concat, "CO")
If Trim(mBQE_ZADRESS0.ADRESSPAY) = "FRANCE" Then mBQE_ZADRESS0.ADRESSPAY = ""
mBEN_BQE = Trim(mBQE_ZADRESS0.ADRESSRA1) & " " & Trim(mBQE_ZADRESS0.ADRESSRA2)

'DR 01/06/2017
'Call rsZENCTIE_Adresse(Left(xZENCCAR0.ENCCARORD, 1), Mid(xZENCCAR0.ENCCARORD, 2), X, xYENCTIE0, mDON_ZADRESS0, mDON_Concat, "CD")
Call rsZENCTIE_Adresse(Left(xZENCCAR0.ENCCARORD, 1), Mid(xZENCCAR0.ENCCARORD, 2), X, xYENCTIE0, mDON_ZADRESS0, mDON_Concat, "CO")
If Trim(mDON_ZADRESS0.ADRESSPAY) = "FRANCE" Then mDON_ZADRESS0.ADRESSPAY = ""

'DR 01/06/2017
'Call rsZENCTIE_Adresse(Left(xZENCCAR0.ENCCARTIR, 1), Mid(xZENCCAR0.ENCCARTIR, 2), X, xYENCTIE0, mBEN_ZADRESS0, mBEN_Concat, "CD")
Call rsZENCTIE_Adresse(Left(xZENCCAR0.ENCCARTIR, 1), Mid(xZENCCAR0.ENCCARTIR, 2), X, xYENCTIE0, mBEN_ZADRESS0, mBEN_Concat, "CO")
If Trim(mBEN_ZADRESS0.ADRESSPAY) = "FRANCE" Then mBEN_ZADRESS0.ADRESSPAY = ""
'_____________________________________________________________________________________
For K = 1 To arrFields_BIA_Nb
    arrFields_BIA_Value(K) = ""
    blnFields_BIA_Name(K) = False
    If Mid$(arrFields_BIA_Lib(K), 1, 1) = Asc34 Then
        K2 = InStr(2, arrFields_BIA_Lib(K), Asc34)
        If K2 = 0 Then K2 = Len(arrFields_BIA_Lib(K))
        arrFields_BIA_Value(K) = Mid$(arrFields_BIA_Lib(K), 2, K2 - 2)
    End If
        
Next K

mDescription = ""
mIrrégularités = ""
mAR_Accord = "accord"
mAR_Courrier = "courrier"
mATTN = "A l'attention de M."
newdocumentRDE = ""
Call Word_RDE(newdocumentRDE)
If Trim(newdocumentRDE) <> "" Then
    Call WebBrowser1.Navigate(newdocumentRDE)
    WebBrowser1.Visible = True
    DoEvents
End If

End Sub

Public Sub paramREMDOC_Init()
Dim X As String

    X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'REMDOC'" _
         & " and BIATABK1 = 'FAX' and BIATABK2 ='" & usrName_UCase & "'"
    Set rsSab = cnsab.Execute(X)
    If rsSab.EOF Then
        X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'REMDOC'" _
             & " and BIATABK1 = 'FAX' and BIATABK2 ='*'"
        Set rsSab = cnsab.Execute(X)
    End If
    If rsSab.EOF Then
        paramREMDOC_FAX = ""
    Else
        paramREMDOC_FAX = Trim(rsSab("BIATABTXT"))
    End If
    X = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'REMDOC'" _
         & " and BIATABK1 = 'TEL_NEGO' and BIATABK2 ='" & usrName_UCase & "'"
    Set rsSab = cnsab.Execute(X)
    If rsSab.EOF Then
        X = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'REMDOC'" _
             & " and BIATABK1 = 'TEL_NEGO' and BIATABK2 ='*'"
        Set rsSab = cnsab.Execute(X)
    End If
    If rsSab.EOF Then
         paramREMDOC_TEL_NEGO = ""
    Else
       paramREMDOC_TEL_NEGO = Trim(rsSab("BIATABTXT"))
    End If

End Sub

Private Function retourne_Banque_Beneficiaire(zBIC As String) As String
Dim X As String
Dim rs As ADODB.Recordset

    retourne_Banque_Beneficiaire = zBIC
    X = "select SWIBICIN1, SWIBICIN2 from " & paramIBM_Library_SAB & ".ZSWIBIC0" _
        & " where SWIBICBIC = '" & zBIC & "'"
    Set rs = cnsab.Execute(X)
    Do While Not rs.EOF
        retourne_Banque_Beneficiaire = Trim(rs("SWIBICIN1")) & Trim(rs("SWIBICIN2"))
        Exit Do
    Loop
    rs.Close
    Set rs = Nothing
    
End Function

Public Function retourne_date_valeur(numDos As Long) As String
Dim xSql As String
Dim rs As ADODB.Recordset

    retourne_date_valeur = ""
    xSql = "select ENCREGDVA from " & paramIBM_Library_SAB & ".ZENCREG0" _
    & " where ENCREGDOS  = " & numDos & " and ENCREGSEN='C'"
    Set rs = cnsab.Execute(xSql)
    Do While Not rs.EOF
        If IsNumeric(Trim(rs("ENCREGDVA"))) Then
            retourne_date_valeur = dateImp10_S(CStr(CLng(Trim(rs("ENCREGDVA"))) + 19000000))
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

End Function

Private Function retourne_TVANIF(z1 As String, z2 As String) As String
Dim X As String

    retourne_TVANIF = ""
    If z1 = "T" Then
        X = "select TVANIFCLIT  from " & paramIBM_Library_SABSPE & ".YTVANIF0 where TVANIFCLI ='" & z2 & "'"
        Set rsSabX = cnsab.Execute(X)
        If Not rsSabX.EOF Then retourne_TVANIF = Trim(rsSabX("TVANIFCLIT"))
    Else
        X = "select CLIFISNIF  from " & paramIBM_Library_SAB & ".ZCLIFIS0" _
              & " where CLIFISETA = 1 and CLIFISCLI ='" & z2 & "' and CLIFISTYP = 1"
        Set rsSabX = cnsab.Execute(X)
        If Not rsSabX.EOF Then retourne_TVANIF = Trim(rsSabX("CLIFISNIF"))
    End If
    
End Function

Private Sub Word_GetStyle(par As Word.Range, debut As Long, ByRef dColor As Long, ByRef dFontname As String, ByRef dItalic As Boolean, ByRef dBold As Boolean, ByRef dSize As Long, ByRef dShading As Long)

    dColor = par.Characters(debut).Font.Color
    dFontname = par.Characters(debut).Font.Name
    dItalic = par.Characters(debut).Font.Italic
    dBold = par.Characters(debut).Font.Bold
    dSize = par.Characters(debut).Font.Size
    dShading = par.Characters(debut).HighlightColorIndex

End Sub

Public Sub lstCourrier_Load()
Dim X As String, K As Integer, K1 As Integer, blnOk As Boolean, blnDisplay As Boolean
Dim wCourrier_Des_Len As Integer, blnDisplay_All As Boolean
Static blnOrderBy As Boolean

On Error Resume Next
If blnOrderBy <> optCourrier_All Then
    blnOrderBy = optCourrier_All
    arrCourrier_Doc_Nb = 0
End If
If arrCourrier_Doc_Nb = 0 Then
    mCourrier_Des_Len = 8
    X = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
       & " where BIATABID = 'REMDOC' and BIATABK1 = 'Courrier_Des'"
    Set rsSab = cnsab.Execute(X)
    If Not rsSab.EOF Then
        ReDim arrCourrier_Doc(rsSab(0) + 1), arrCourrier_Des(rsSab(0) + 1), arrCourrier_Id(rsSab(0) + 1) _
            , arrCourrier_Originaux_Param_Nb(rsSab(0) + 1), arrCourrier_Originaux_Dossier_Nb(rsSab(0) + 1)
    End If
    X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 A , " & paramIBM_Library_SABSPE & ".YBIATAB0 B" _
       & " where A.BIATABID = 'REMDOC' and A.BIATABK1 = 'Courrier_Des'" _
       & " and B.BIATABID = 'REMDOC' and B.BIATABK1 = 'Courrier_Doc'" _
       & " and A.BIATABK2 = B.BIATABK2"
    If optCourrier_All.value = True Then
        X = X & " order by B.BIATABTXT"
    Else
        X = X & " order by substring(A.BIaTABTXT , 124 , 5)"
    End If
    Set rsSab = cnsab.Execute(X)
    Do While Not rsSab.EOF
        arrCourrier_Doc_Nb = arrCourrier_Doc_Nb + 1
        arrCourrier_Des(arrCourrier_Doc_Nb) = rsSab(3)
        arrCourrier_Doc(arrCourrier_Doc_Nb) = Trim(rsSab(7))
        arrCourrier_Id(arrCourrier_Doc_Nb) = Trim(rsSab(2))
        arrCourrier_Originaux_Param_Nb(arrCourrier_Doc_Nb) = Val(Mid$(rsSab(3), 122, 1))
        If arrCourrier_Originaux_Param_Nb(arrCourrier_Doc_Nb) = 0 Then arrCourrier_Originaux_Param_Nb(arrCourrier_Doc_Nb) = 1
        rsSab.MoveNext
    Loop
End If
For K = 0 To arrCourrier_Doc_Nb
    arrCourrier_Originaux_Dossier_Nb(K) = arrCourrier_Originaux_Param_Nb(K)
Next K
wCourrier_Des_Len = mCourrier_Des_Len
blnDisplay_All = False
mCourrier_Des = String(wCourrier_Des_Len, " ")
If xZENCCAR0.ENCCARCOP = "RDE" Then
    Mid$(mCourrier_Des, 1, 1) = "O"
    Mid$(mCourrier_Des, 2, 1) = ""
Else
    Mid$(mCourrier_Des, 2, 1) = "O"
    Mid$(mCourrier_Des, 1, 1) = ""
End If
If optLangue_FR.value = True Then
    Mid$(mCourrier_Des, 3, 1) = "O"
    Mid$(mCourrier_Des, 4, 1) = ""
End If
If optLangue_GB.value = True Then
    Mid$(mCourrier_Des, 4, 1) = "O"
    Mid$(mCourrier_Des, 3, 1) = ""
End If
If optCourrier_OUV.value = True Then
    Mid$(mCourrier_Des, 6, 1) = "O"
    Mid$(mCourrier_Des, 7, 1) = ""
    Mid$(mCourrier_Des, 8, 1) = ""
ElseIf optCourrier_REG.value = True Then
    Mid$(mCourrier_Des, 8, 1) = "O"
    Mid$(mCourrier_Des, 6, 1) = ""
    Mid$(mCourrier_Des, 7, 1) = ""
Else
    Mid$(mCourrier_Des, 6, 1) = ""
    Mid$(mCourrier_Des, 7, 1) = ""
    Mid$(mCourrier_Des, 8, 1) = ""
End If
If optCourrier_All.value = True Then wCourrier_Des_Len = 4: blnDisplay_All = True
lstCourrier.Clear
X = ""
For K = 1 To arrCourrier_Doc_Nb
    blnOk = True: blnDisplay = True
    If mCourrier_Des <> Mid$(arrCourrier_Des(K), 1, wCourrier_Des_Len) Then
        For K1 = 1 To wCourrier_Des_Len
            Select Case Mid$(arrCourrier_Des(K), K1, 1)
                Case "O"
                    Select Case Mid$(mCourrier_Des, K1, 1)
                        Case "O"
                        Case Else: blnOk = False: blnDisplay = False: Exit For
                    End Select
                 Case "N"
                    Select Case Mid$(mCourrier_Des, K1, 1)
                        Case "N":
                        Case Else: blnOk = False: blnDisplay = False: Exit For
                    End Select
            End Select
        Next K1
        If Mid$(arrCourrier_Des(K), 123, 1) = "O" Then blnOk = False
        If blnDisplay_All Then blnOk = False
    End If
    If Mid$(arrCourrier_Des(K), 124, 1) = "O" Then
        X = "+ " & arrCourrier_Doc(K)
    Else
        X = "  " & arrCourrier_Doc(K)
    End If
    If blnDisplay Then
        If blnOk Then
            lstCourrier.AddItem X
            lstCourrier.Selected(lstCourrier.ListCount - 1) = True
        Else
            lstCourrier.AddItem X
        End If
    End If
Next K
lstCourrier.Visible = arrHab(2)
End Sub

Private Sub Word_SetStyle(ByRef par As Word.Range, debut As Long, longueur As Long, dColor As Long, dFontname As String, dItalic As Boolean, dBold As Boolean, dSize As Long, dShading As Long)
Dim ii As Long
Dim rng As Word.Range

    For ii = 0 To longueur - 1
        If (ii + debut) <= Len(par) Then
            par.Characters(ii + debut).Font.Color = dColor
            par.Characters(ii + debut).Font.Name = dFontname
            par.Characters(ii + debut).Font.Italic = dItalic
            par.Characters(ii + debut).Font.Bold = dBold
            par.Characters(ii + debut).Font.Size = dSize
            If dShading <> 0 Then
                par.Characters(ii + debut).HighlightColorIndex = dShading
            End If
        Else
            Exit For
        End If
    Next ii

End Sub



Private Sub cmdPrint_Courrier_Init()
Dim X As String, K As Integer, K1 As Integer, K_Originaux_Nb As Integer
Dim wCourrier_Id As Long

On Error GoTo Error_Handler

Call lstErr_Clear(lstErr, cmdContext, "Courrier_Init ...."): DoEvents
lstErr.Height = 510

mClipBoard = Clipboard.GetText
Clipboard.Clear

fraInfo_M.Visible = False
fgInfo_M.Visible = False
txtInfo_M.Visible = False
cmdInfo_M_Ok.Visible = False
cmdPrint.Visible = False

fgInfo_M_Reset

ReDim arrDOC(lstCourrier.ListCount * 3 + 1), arrDOC_FileName(lstCourrier.ListCount * 3 + 1) _
    , arrDOC_REF(lstCourrier.ListCount * 3 + 1), arrDoc_Originaux_Nb(lstCourrier.ListCount * 3 + 1)

arrDoc_Nb = 0
mANNEXES_NB = 0
mDoc_Page_Nb = 0

For K_Originaux_Nb = 1 To 3
    For K = 0 To lstCourrier.ListCount - 1
        If lstCourrier.Selected(K) Then
            lstCourrier.ListIndex = K
            
            X = RTrim(lstCourrier.Text)
            X = Mid$(X, 3, Len(X) - 2)
            wCourrier_Id = 0
            
            For K1 = 1 To arrCourrier_Doc_Nb
                If X = arrCourrier_Doc(K1) Then
                    wCourrier_Id = Val(arrCourrier_Id(K1))
                    Exit For
                End If
            Next K1
            
            If wCourrier_Id > 0 Then
                If arrCourrier_Originaux_Dossier_Nb(K1) >= K_Originaux_Nb Then
                    arrDoc_Nb = arrDoc_Nb + 1
                    arrDOC_REF(arrDoc_Nb) = wCourrier_Id
                    arrDOC(arrDoc_Nb) = paramRDE_Dossier_Path_DROPI & "Modèles\" & X
                    If arrDoc_Nb = 1 Then mDOS_Modèle = X
                    arrDOC_FileName(arrDoc_Nb) = X
                    If K_Originaux_Nb = 1 And Mid$(lstCourrier.Text, 1, 1) = "+" Then mANNEXES_NB = mANNEXES_NB + 1
                    If K_Originaux_Nb = 1 Then mDoc_Page_Nb = mDoc_Page_Nb + 1
                End If
            End If
        End If
    
    Next K
Next K_Originaux_Nb


If arrDoc_Nb = 0 Then
    Call MsgBox("Choisissez au moins un document")
    Exit Sub
End If
'__________________________________________________________________

ProgressBar1.Visible = True
ProgressBar1.Min = 0: ProgressBar1.Max = 10
ProgressBar1.value = 1

Set appWord = New Word.Application

appWord.Visible = False

ProgressBar1.value = ProgressBar1.value + 1

hwndWord = FindWindow(vbNullString, "Microsoft Word")

mWord_PDF_Path = Environ("commonprogramfiles") & "\Microsoft Shared\OFFICE" _
& Format(Val(appWord.Version), "00") & "\EXP_PDF.DLL"

'__________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Lecture des modèles  : " & arrDoc_Nb): DoEvents
ProgressBar1.value = ProgressBar1.value + 1

docWord_Concatenate appWord, arrDOC(), arrDoc_Nb, arrDOC_REF()
cmdPrint_Courrier_Word_Tables

Call lstErr_AddItem(lstErr, cmdContext, "Recherche variable '?BIA' ...."): DoEvents
ProgressBar1.value = ProgressBar1.value + 1
appWord.Selection.WholeStory
With appWord.Selection.Find
    .Text = "?"
    .Execute
End With
If appWord.Selection.Find.Found Then
    Call cmdPrint_Courrier_Info_M
Else
    blnPrint_Courrier_Ok = True
End If
ProgressBar1.value = ProgressBar1.value + 1
If blnPrint_Courrier_Ok Then Call cmdPrint_Courrier_Word

ProgressBar1.Visible = False
    
    GoTo Exit_sub
Error_Handler:
    MsgBox Error
    cmdPrint_Courrier_Word_Quit
Exit_sub:

End Sub
Private Sub cmdPrint_Dossier_Init()
Dim X As String, wFile As String
On Error GoTo Error_Handler

Call lstErr_Clear(lstErr, cmdContext, "Courrier_Init ...."): DoEvents
lstErr.Height = 510

wFile = "C:\Temp\" & mCOP_DOS & ".rtf"
If Dir(wFile) <> "" Then Kill wFile

ReDim arrDOC(2), arrDOC_REF(2)

arrDoc_Nb = 1
arrDOC(arrDoc_Nb) = wFile
'__________________________________________________________________

ProgressBar1.Visible = True
ProgressBar1.Min = 0: ProgressBar1.Max = 10
ProgressBar1.value = 1

Set appWord = New Word.Application
ProgressBar1.value = ProgressBar1.value + 1

mWord_PDF_Path = Environ("commonprogramfiles") & "\Microsoft Shared\OFFICE" _
& Format(Val(appWord.Version), "00") & "\EXP_PDF.DLL"

'__________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Lecture : " & wFile): DoEvents
ProgressBar1.value = ProgressBar1.value + 1
docWord_Concatenate appWord, arrDOC(), arrDoc_Nb, arrDOC_REF()

'blnPrint_Courrier_Ok = True

ProgressBar1.value = ProgressBar1.value + 1
Call lstErr_AddItem(lstErr, cmdContext, "Enregistrement .pdf ...."): DoEvents
mDOS_seq = mDOS_seq + 1
wFile = "C:\Temp\" & mCOP_DOS & "_" & DSYS_Time & mDOS_seq & ".pdf"
    
Call appWord.ActiveDocument.ExportAsFixedFormat(wFile, wdExportFormatPDF, False, wdExportOptimizeForPrint)
            
ProgressBar1.Visible = False
     cmdPrint_Courrier_Word_Quit
   
    GoTo Exit_sub
Error_Handler:
    MsgBox Error
    cmdPrint_Courrier_Word_Quit
Exit_sub:

End Sub







Sub docWord_Concatenate(appWord As Word.Application, arrDOC() As String, lDOC_NB As Integer, arrDOC_REF() As String)
Dim NewDoc As Boolean
Dim current As String
Dim NewDocName As String
Dim K As Integer

    NewDoc = False
    For K = 1 To lDOC_NB
                ' suppression temporaire de l'update automatique des links (évite l'apparition d'un warning message à chaque ouverture d'un fichier doc)
                ' temporary delete of links' automatic update (in order to avoid the appearance of a warning message each time a .doc file is opened)
                appWord.Options.UpdateLinksAtOpen = False
                ' ouverture du fichier sans le rendre visible
                ' opening the file without making him visible
                appWord.Documents.Open FileName:=arrDOC(K), Visible:=True, _
                ConfirmConversions:=False, ReadOnly:=False, AddToRecentFiles:=False, _
                PasswordDocument:="", PasswordTemplate:="", Revert:=False, _
                WritePasswordDocument:="", WritePasswordTemplate:="", Format:= _
                wdOpenFormatAuto
                current = appWord.ActiveDocument.Name
                ' sélection de l'ensemble des données du document
                ' selecting all datas from the document
                If Not NewDoc Then 'Je ne crée un nouveau document au modèle que s'il n'existe pas
                    appWord.Documents.Add
                    NewDocName = appWord.ActiveDocument.Name
                    NewDoc = True ' On ne passe ici qu'une fois
                End If
                'Retour dans le document à copier
                appWord.Documents(current).Activate
                'Sélection de tout le document à copier
                appWord.Selection.WholeStory
                       With appWord.Selection.Find
                        .Text = "#BIA_DOC_REF"
                        .Replacement.Text = arrDOC_REF(K)
                        .Execute Replace:=wdReplaceAll
                     End With
                ' copie de toutes les données
                ' copy of all datas
                appWord.Selection.Copy
                'Retour dans le nouveau document
                appWord.Documents(NewDocName).Activate
                ' colle toute les données sauvegardées dans le document compilé
                ' pasting all saved datas in the compiled document
                appWord.Selection.PasteAndFormat wdPasteDefault
                ' fermeture du document (copié) sans sauvegarde
                ' closing  the document without saving it
                appWord.Documents(current).Close (wdDoNotSaveChanges)
                'On va en fin du doc créé pour être en position de recevoir la nouvelle copie
                appWord.Selection.EndKey Unit:=wdLine, Extend:=wdMove
                If K < lDOC_NB Then appWord.Selection.InsertBreak Type:=wdPageBreak
                ' réactivation de l'option update automatique des liens
                ' reactivating links' automatic update option
                appWord.Options.UpdateLinksAtOpen = True
        Next K
    appWord.ActiveDocument.Content.Select
    appWord.Selection.Fields.Update
appWord.Documents(NewDocName).Activate
End Sub

Sub docWord_Filigrane(appWord As Word.Application, Texte As String, lColor As Long)
    Dim Section As Word.Section
    Dim Header As Word.HeaderFooter
    appWord.ActiveDocument.Content.Select

    appWord.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    For Each Section In appWord.ActiveDocument.Sections
        For Each Header In Section.Headers
            docWord_Filigrane_Add appWord, Texte, Header, Section, lColor
        Next
    Next
    appWord.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub


Private Sub docWord_Filigrane_Add(appWord As Word.Application, Texte As String, Header As Word.HeaderFooter, Section As Word.Section, lColor As Long)
    Dim ShapeName As String
    Dim Shape
    ShapeName = "Filigrane_" & Section.Index & "_" & Header.Index
    Header.Range.Select
    
    'détruit un éventuel Filigrane précédent
    On Error Resume Next
    Set Shape = Header.Shapes(ShapeName)
    If Not Shape Is Nothing Then Shape.Delete
    If Texte = "" Then Exit Sub
    On Error GoTo Error_Handler
    'ajoute le Filigrane (c'est dans l'entete, et ça prend 1x1 point en haut à gauche de la page)
    Set Shape = appWord.Selection.HeaderFooter.Shapes.AddTextEffect( _
        Office.MsoPresetTextEffect.msoTextEffect1, _
        Texte, "Calibri", 1, False, False, _
        0, 0)
        
    Shape.Select
    'met en forme le Filigrane pour prendre toute la page
    With appWord.Selection.ShapeRange
        .Name = ShapeName
        .TextEffect.Text = Texte
        .TextEffect.FontName = "Calibri"
        .TextEffect.FontSize = 1 'la taille de la police est fixé par le ratio
        .Line.Visible = False
        .Fill.Visible = True
        .Fill.Solid
        .Fill.ForeColor.RGB = lColor ' WdColor.wdColorLightGreen ' .wdColorRed
        .Fill.Transparency = 0.8
        .Rotation = 305
        .LockAspectRatio = True
        .Height = appWord.CentimetersToPoints(3.22)
        .Width = appWord.CentimetersToPoints(19.34)
        .WrapFormat.AllowOverlap = True
        .WrapFormat.Side = wdWrapNone
        .WrapFormat.Type = 3
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .RelativeVerticalPosition = wdRelativeHorizontalPositionPage
        .Left = wdShapeCenter
        .Top = wdShapeCenter
    End With
    Set Shape = Nothing
    Exit Sub
Error_Handler:
    Call MsgBox(Error, vbCritical, "docWord_Filigrane_Add")
End Sub








Public Function Word_ZENCCOM0(tblpar As Word.Table) As Long
Dim X As String
Dim c_com() As String
Dim lib_com() As String
Dim t_com() As String
Dim m_com() As Currency
Dim tva_com() As Currency
Dim maxI As Long
Dim I As Long
Dim ii As Long

    X = "select count(ENCCOMDOS) as MAXI  from " & paramIBM_Library_SAB & ".ZENCCOM0" _
          & " where ENCCOMETA = " & xZENCCAR0.ENCCARETA & " and ENCCOMAGE = " & xZENCCAR0.ENCCARAGE _
          & " and ENCCOMSER = '" & xZENCCAR0.ENCCARSER & "' and ENCCOMSSE = '" & xZENCCAR0.ENCCARSSE & "'" _
          & " and ENCCOMCOP = '" & xZENCCAR0.ENCCARCOP & "' and ENCCOMDOS = " & xZENCCAR0.ENCCARDOS
    Set rsSab = cnsab.Execute(X)
    maxI = 0
    Do While Not rsSab.EOF
        maxI = CLng(rsSab("MAXI"))
        Exit Do
    Loop
    mMTD_NET = CDbl(xZENCCAR0.ENCCARMON)
    mTVA_T = 0
    mTOTAL_TTC = 0
    If maxI > 6 Then maxI = 6 'on ne gère que 6 commissions maximum
    If maxI = 0 Then
        'suppression du texte de la 1ère ligne
        For ii = 1 To 4
            tblpar.Cell(38, ii).Range.Text = ""
        Next ii
        Word_ZENCCOM0 = 37 'la ligne d'entete du tableau
        Exit Function
    End If
    ReDim c_com(1 To maxI)
    ReDim lib_com(1 To maxI)
    ReDim t_com(1 To maxI)
    ReDim m_com(1 To maxI)
    ReDim tva_com(1 To maxI)
    X = "select ENCCOMCOM,ENCCOMTPC,ENCCOMMCD,ENCCOMMTD,ENCCOMSEQ,BASTABDON" _
        & " from SAB073.ZENCCOM0, SAB073.ZBASTAB0" _
        & " where ENCCOMETA = " & xZENCCAR0.ENCCARETA & " and ENCCOMAGE = " & xZENCCAR0.ENCCARAGE _
        & " and ENCCOMSER = '" & xZENCCAR0.ENCCARSER & "' and ENCCOMSSE = '" & xZENCCAR0.ENCCARSSE & "'" _
        & " and ENCCOMCOP = '" & xZENCCAR0.ENCCARCOP & "' and ENCCOMDOS = " & xZENCCAR0.ENCCARDOS _
        & " and BASTABNUM=44 and BASTABARG=ENCCOMCOM" _
        & " ORDER BY ENCCOMSEQ"
    Set rsSab = cnsab.Execute(X)
    maxI = 0
    Do While Not rsSab.EOF
        maxI = maxI + 1
        If maxI > 6 Then Exit Do
        c_com(maxI) = rsSab("ENCCOMCOM")
        lib_com(maxI) = Trim(Left(rsSab("BASTABDON"), 30))
        t_com(maxI) = rsSab("ENCCOMTPC")
        m_com(maxI) = CDbl(rsSab("ENCCOMMCD"))
        tva_com(maxI) = CDbl(rsSab("ENCCOMMTD"))
        rsSab.MoveNext
    Loop
    For I = 1 To maxI
        tblpar.Cell(37 + I, 1).Range.Text = c_com(I) & " " & lib_com(I)
        tblpar.Cell(37 + I, 2).Range.Text = t_com(I)
        If m_com(I) <> 0 Then
            tblpar.Cell(37 + I, 3).Range.Text = Format(m_com(I), "### ### ##0.00")
        Else
            tblpar.Cell(37 + I, 3).Range.Text = ""
        End If
        If tva_com(I) <> 0 Then
            tblpar.Cell(37 + I, 4).Range.Text = Format(tva_com(I), "### ##0.00")
        Else
            tblpar.Cell(37 + I, 4).Range.Text = ""
        End If
        mMTD_NET = mMTD_NET - m_com(I) - tva_com(I)
        mTOTAL_TTC = mTOTAL_TTC + m_com(I) + tva_com(I)
        mTVA_T = mTVA_T + tva_com(I)
        If m_com(I) <> 0 Then
            If InStr(LCase(lib_com(I)), "frais du correspondant") > 0 Then
                mFRAIS2 = m_com(I)
            ElseIf InStr(LCase(lib_com(I)), "frais") > 0 Then
                mFRAIS1 = m_com(I)
            Else
                mCOM1 = m_com(I)
            End If
        End If
    Next I
    Word_ZENCCOM0 = 37 + maxI
    
End Function





Public Sub Word_RDE(ByRef newdocumentRDE As String)
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oTable As Word.Table
Dim oPar As Word.Paragraph
Dim oTbl As Word.Table
Dim K As Integer
Dim X As String
Dim ou As Long
Dim ii As Long

    If arrHab(2) Then
        K = Windows_Processus_Actif("WINWORD")
        If K > 0 Then
               Call MsgBox("Attention, il y a déjà " & K & " instance(s) 'Word' active(s)!" & vbCrLf & vbCrLf & "Veuillez fermer les documents 'Word', si vous devez éditer des courriers", vbExclamation, "SAB_Dossier : courrier")
        End If
    End If
    
    Set oWord = CreateObject("Word.Application")
    oWord.Visible = False
    Set oDoc = oWord.Documents.Open(paramRDE_Dossier_Path_DROPI & "Modèles\" & "ZENCCAR0.docx", False, False)
    DoEvents
    Set oTbl = oDoc.Tables(1)
    oTbl.Cell(1, 2).Range.Text = mCOP_DOS
    oTbl.Cell(1, 3).Range.Text = mETA
    If xZENCCAR0.ENCCARDAR = 0 Then
        oTbl.Cell(2, 2).Range.Text = ""
    Else
        oTbl.Cell(2, 2).Range.Text = dateImp10_S(xZENCCAR0.ENCCARDAR + 19000000)
    End If
    If xZENCCAR0.ENCCARDAM = 0 Then
        oTbl.Cell(2, 4).Range.Text = ""
    Else
        oTbl.Cell(2, 4).Range.Text = dateImp10_S(xZENCCAR0.ENCCARDAM + 19000000)
    End If
    If xZENCCAR0.ENCCARDAC = 0 Then
        oTbl.Cell(2, 6).Range.Text = ""
    Else
        oTbl.Cell(2, 6).Range.Text = dateImp10_S(xZENCCAR0.ENCCARDAC + 19000000)
    End If
    oTbl.Cell(3, 2).Range.Text = xZENCCAR0.ENCCARREX
    oTbl.Cell(4, 2).Range.Text = xZENCCAR0.ENCCARRCL
    oTbl.Cell(5, 2).Range.Text = xZENCCAR0.ENCCARTYR
    oTbl.Cell(6, 2).Range.Text = xZENCCAR0.ENCCARNAT
    oTbl.Cell(8, 2).Range.Text = mMON_DEV
    oTbl.Cell(9, 2).Range.Text = xZENCCAR0.ENCCARAUT
    w_ZADRESSE0 = mBQE_ZADRESS0
    Call Word_ZENCCAR0_ADR(oTbl, "#BQE_RS")
    w_ZADRESSE0 = mDON_ZADRESS0
    Call Word_ZENCCAR0_ADR(oTbl, "#DON_RS")
    w_ZADRESSE0 = mBEN_ZADRESS0
    Call Word_ZENCCAR0_ADR(oTbl, "#BEN_RS")
    ou = Word_ZENCCOM0(oTbl)
    If ou = 37 Then
        ou = 45
    Else
        ou = ou + 3
    End If
    Call Word_MT499(oTbl, ou)
    Set oTbl = Nothing
    newdocumentRDE = paramTemp_Folder & "\ZENCCAR0_" & Format(Now, "hh_nn_ss") & ".htm"
    oDoc.SaveAs newdocumentRDE, wdFormatHTML
    oDoc.Close
    Set oDoc = Nothing
    oWord.Quit
    Set oWord = Nothing
    
End Sub
Private Sub Word_MT499(tblpar As Word.Table, ou As Long)
Dim txtMT499 As String
Dim mSWISABSWID_499 As Long

Dim X As String, xSql As String
Dim xValue As String, V
Dim K As Integer, K2 As Integer, iAsc13 As Integer, iLen As Integer
Dim xField As String
On Error GoTo Error_Handler

If Not blnSIDE_DB_Open Then
    cnSIDE_DB.Open paramODBC_DSN_SIDE_DB
    blnSIDE_DB_Open = True
End If

txtMT499 = ""
tblpar.Cell(ou, 1).Range.Text = ""
mSWISABSWID_499 = 0
'__________________________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABOPEC = '" & xZENCCAR0.ENCCARCOP & "'" _
     & " and   SWISABOPEN = " & xZENCCAR0.ENCCARDOS _
     & " and  SWISABWMTK = '499' and SWISABWES = 'E' order by SWISABSWID"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    mSWISABSWID_499 = rsSab("SWISABSWID")
    txtMT499 = "MT" & rsSab("SWISABWMTK") & " de " & rsSab("SWISABWBIC") _
         & " reçu le " & dateImp10_S(rsSab("SWISABWAMJ")) & " " & timeImp8(rsSab("SWISABWHMS")) _
         & "    (" & Trim(rsSab("SWISABSWID")) & ")"
    
    tblpar.Cell(ou, 1).Range.Text = txtMT499
    txtMT499 = ""
    Call arrMT_Fields_Load(rsSab("SWISABWMTK"))
    xSql = "select *  from rtextField  " _
        & "where Aid = " & rsSab("SWISABWID1") _
        & " and text_s_umidl = " & rsSab("SWISABWIDL") _
        & " and text_s_umidh  =  " & rsSab("SWISABWIDH") _
        & " order by field_cnt"
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    If Not rsSIDE_DB.EOF Then
        Do While Not rsSIDE_DB.EOF
            Select Case rsSIDE_DB("field_code")
                Case "45", "46", "47", "77":
                    V = rsSIDE_DB("value_memo")
                    If IsNull(V) Then V = rsSIDE_DB("value")
                Case Else:
                        V = rsSIDE_DB("value")
            End Select
            If IsNull(V) Then
                xValue = ""
            Else
                xValue = V
            End If
            xField = rsSIDE_DB("field_code") & rsSIDE_DB("field_option")
            If IsNumeric(xField) Then
                txtMT499 = txtMT499 & vbCrLf & xField & " : " & arrMT_Fields_Scan(xField)
            Else
                txtMT499 = txtMT499 & vbCrLf & xField & " " & arrMT_Fields_Scan(xField)
            End If
                iLen = Len(xValue)
                K = 1
                Do
                   iAsc13 = InStr(K, xValue, Asc13)
                   If iAsc13 > 0 Then
                        txtMT499 = txtMT499 & vbCrLf & Trim(Mid$(xValue, K, iAsc13 - K))
                       K = iAsc13 + 2
                   End If
                Loop Until iAsc13 = 0
                txtMT499 = txtMT499 & vbCrLf & Trim(Mid$(xValue, K, iLen - K + 1))
            rsSIDE_DB.MoveNext
        Loop
    Else
'________________________________________________________________________________
        xSql = "select * from rtext " _
            & "where Aid = " & rsSab("SWISABWID1") _
            & " and text_s_umidl = " & rsSab("SWISABWIDL") _
            & " and text_s_umidh  =  " & rsSab("SWISABWIDH")
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
        If Not rsSIDE_DB.EOF Then
            Call srvrText_GetBuffer_ODBC(rsSIDE_DB, xrText)
            xValue = xrText.text_data_block & Asc13
            iLen = Len(xValue)
            If Mid$(xValue, 1, 3) = Asc13 & Asc10 & ": " Then
                K = 3
            Else
                K = 1
            End If
            Do
                iAsc13 = InStr(K, xValue, Asc13)
                If iAsc13 > 0 Then
                    X = Trim(Mid$(xValue, K, iAsc13 - K))
                    If Mid$(X, 1, 1) <> ":" Then
                        txtMT499 = txtMT499 & vbCrLf & Trim(Mid$(xValue, K, iAsc13 - K))
                    Else
                        K2 = InStr(2, X, ":")
                        If K2 > 0 Then
                            xField = Mid$(X, 2, K2 - 2)
                            txtMT499 = txtMT499 & vbCrLf & Trim(Mid$(X, 2, K2 - 1)) & " " & arrMT_Fields_Scan(xField)
                            X = Trim(Mid$(X, K2 + 1, Len(X) - K2))
                             txtMT499 = txtMT499 & X
                                Select Case xField
                                    Case "52A", "57A", "57D", "58A", "59A", "59F": txtMT499 = txtMT499 & retourne_Banque_Beneficiaire(X)
                                End Select
                        Else
                            txtMT499 = txtMT499 & vbCrLf & Trim(Mid$(xValue, K, iAsc13 - K))
                        End If
                    End If
                    K = iAsc13 + 2
                End If
             Loop Until iAsc13 = 0
        End If
    End If
    txtMT499 = txtMT499 & vbCrLf
    rsSab.MoveNext
Loop

tblpar.Cell(ou + 1, 1).Range.Text = txtMT499
'_______________________________________________________________________

GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbCritical, currentAction
Exit_sub:

End Sub

Private Sub Word_Traite_Replace(par As Word.Paragraph, src As String, dest As String, insertion As Long)
Dim ii As Long
Dim z As String
Dim fSize As Long
Dim fcolor As Long
Dim fBold As Boolean
Dim fShading As Long
Dim oRng As Word.Range
Dim fItalic As Boolean
Dim ffontname As String

    z = Trim(dest)
    Set oRng = par.Range
    Call Word_GetStyle(oRng, insertion, fcolor, ffontname, fItalic, fBold, fSize, fShading)
    For ii = 1 To Len(src)
        oRng.Characters(insertion).Delete
    Next ii
    If insertion = 1 Then insertion = 2
    If Asc(oRng.Characters(insertion - 1).Text) = 13 Then
        oRng.Characters(insertion - 1).insertBefore (z)
    Else
        oRng.Characters(insertion - 1).InsertAfter (z)
    End If
    Call Word_SetStyle(oRng, insertion, Len(z), fcolor, ffontname, fItalic, fBold, fSize, fShading)
    Set oRng = Nothing
    
End Sub

Public Sub Word_ZENCCAR0_ADR(tblpar As Word.Table, sujet As String)
Dim X As String
Dim insertion As Long
Dim mDON_TVANIFCLIT As String

    If sujet = "#BQE_RS" Then
        insertion = 10
    ElseIf sujet = "#DON_RS" Then
        insertion = 18
    ElseIf sujet = "#BEN_RS" Then
        insertion = 28
    End If
    X = ""
    If w_ZADRESSE0.ADRESSTYP = "1" Then
        X = Trim(w_ZADRESSE0.ADRESSNUM) & " - "
    Else
        X = w_ZADRESSE0.ADRESSTYP & " " & Trim(w_ZADRESSE0.ADRESSNUM) & " - "
    End If
    If Len(Trim(w_ZADRESSE0.ADRESSRA1)) < Len(w_ZADRESSE0.ADRESSRA1) And Trim(w_ZADRESSE0.ADRESSRA2) <> "" Then
        tblpar.Cell(insertion + 1, 2).Range.Text = X & Trim(w_ZADRESSE0.ADRESSRA1) & Chr(13) & Chr(10) & Trim(w_ZADRESSE0.ADRESSRA2)
    Else
        tblpar.Cell(insertion + 1, 2).Range.Text = X & Trim(w_ZADRESSE0.ADRESSRA1) & Trim(w_ZADRESSE0.ADRESSRA2)
    End If
    If Trim(w_ZADRESSE0.ADRESSAD1) <> "" Then
        tblpar.Cell(insertion + 2, 2).Range.Text = Trim(w_ZADRESSE0.ADRESSAD1)
    End If
    If Trim(w_ZADRESSE0.ADRESSAD2) <> "" Then
        tblpar.Cell(insertion + 3, 2).Range.Text = Trim(w_ZADRESSE0.ADRESSAD2)
    End If
    If Trim(w_ZADRESSE0.ADRESSAD3) <> "" Then
        tblpar.Cell(insertion + 4, 2).Range.Text = Trim(w_ZADRESSE0.ADRESSAD3)
    End If
    X = Trim(w_ZADRESSE0.ADRESSCOP) & " " & Trim(w_ZADRESSE0.ADRESSVIL)
    If Trim(X) <> "" Then
        tblpar.Cell(insertion + 5, 2).Range.Text = Trim(X)
    End If
    If Trim(w_ZADRESSE0.ADRESSPAY) <> "" Then
        tblpar.Cell(insertion + 6, 2).Range.Text = Trim(w_ZADRESSE0.ADRESSPAY)
    End If
    X = ""
    If Trim(w_ZADRESSE0.ADRESSTEL) <> "" Then X = "Tél :  " & Trim(w_ZADRESSE0.ADRESSTEL)
    If Trim(w_ZADRESSE0.ADRESSFAX) <> "" Then X = X & " - Fax :  " & Trim(w_ZADRESSE0.ADRESSFAX)
    If Trim(w_ZADRESSE0.ADRESSTEX) <> "" Then X = X & " - Tlx :  " & Trim(w_ZADRESSE0.ADRESSTEX)
    If Trim(X) <> "" Then
        tblpar.Cell(insertion + 7, 2).Range.Text = Trim(X)
    End If
    X = ""
    If sujet = "#DON_RS" Then
        mDON_TVANIFCLIT = retourne_TVANIF(Left(xZENCCAR0.ENCCARORD, 1), Mid(xZENCCAR0.ENCCARORD, 2))
        If mDON_TVANIFCLIT <> "" Then
            X = "NIF : " & mDON_TVANIFCLIT
            tblpar.Cell(insertion + 8, 2).Range.Text = Trim(X)
            tblpar.Cell(insertion + 8, 2).Range.Font.Color = wdColorGreen
            tblpar.Cell(insertion + 8, 2).Range.Font.Underline = True
        Else
            X = "NIF : ????" & mDON_TVANIFCLIT
            tblpar.Cell(insertion + 8, 2).Range.Text = Trim(X)
            tblpar.Cell(insertion + 8, 2).Range.Font.Underline = False
            tblpar.Cell(insertion + 8, 2).Range.Font.Bold = True
            tblpar.Cell(insertion + 8, 2).Range.Font.Color = wdColorRed
       End If
       X = ""
    End If

End Sub


Private Sub btnControle_Click()
Dim oWord As Word.Application
Dim K As Long
Dim iRow As Long
Dim numDoc As String
Dim oDoc As Word.Document
Dim occurFound() As String
Dim occurTofind() As String
Dim xSql As String
Dim rs As ADODB.Recordset
Dim oKIndice As Long
Dim oK As Boolean
Dim kO As Boolean

    btnControle.Enabled = True
    btnControle.BackColor = ColorConstants.vbRed
    Me.MousePointer = vbHourglass
    lstParam_Modèles_Temp.Clear
    lstParam_Modèles_Temp.ForeColor = ColorConstants.vbRed
    lstParam_Modèles_Temp.FontBold = True
    Set oWord = CreateObject("Word.Application")
    oWord.Visible = False
    For iRow = 0 To lstParam_Modèles_REMDOC.ListCount - 1
        oK = True
        lstParam_Modèles_REMDOC.Selected(iRow) = True
        DoEvents
        numDoc = Retourne_Num_Document("REMDOC", lstParam_Modèles_REMDOC.List(iRow))
        If numDoc <> "" Then
            Set oDoc = oWord.Documents.Open(paramRDE_Dossier_Path_DROPI & "Modèles\" & lstParam_Modèles_REMDOC.List(iRow), False, False)
            DoEvents
            oWord.Selection.WholeStory
            ReDim occurTofind(0 To arrFields_SAB_Nb)
            occurTofind(0) = "0"
            For K = 1 To arrFields_SAB_Nb
                If Trim(arrFields_SAB_Name(K)) <> "#BIA_DOC_REF" Then
                    With oWord.Selection.Find
                        .Forward = True
                        .MatchWholeWord = True
                        .MatchCase = True
                        .Wrap = wdFindContinue
                        .Text = Trim(arrFields_SAB_Name(K))
                        .Execute
                    End With
                    If oWord.Selection.Find.Found Then
                        occurTofind(0) = CStr(Val(occurTofind(0)) + 1)
                        occurTofind(Val(occurTofind(0))) = arrFields_SAB_Name(K)
                    End If
                End If
            Next K
            xSql = "select BIATABK2 from " & paramIBM_Library_SABSPE & ".YBIATAB0"
            xSql = xSql & " where BIATABID = 'REMDOC_#SAB' and BIATABK1 = '" & numDoc & "'"
            Set rs = cnsab.Execute(xSql)
            If rs.EOF Or rs.BOF Then
                If Val(occurTofind(0)) > 0 Then
                    oK = False
                End If
            End If
            If oK Then
                ReDim occurFound(0 To arrFields_SAB_Nb)
                occurFound(0) = "0"
                Do While Not rs.EOF
                    occurFound(0) = CStr(Val(occurFound(0)) + 1)
                    occurFound(Val(occurFound(0))) = Trim(rs("BIATABK2"))
                    rs.MoveNext
                Loop
                rs.Close
                For K = 1 To Val(occurTofind(0))
                    kO = True
                    For oKIndice = 1 To Val(occurFound(0))
                        If Trim(occurFound(oKIndice)) = Trim(occurTofind(K)) Then
                            kO = False
                            Exit For
                        End If
                    Next oKIndice
                    If kO Then
                        oK = False
                        Exit For
                    End If
                Next K
                If Val(occurTofind(0)) <> Val(occurFound(0)) Then oK = False
            End If
            If Not oK Then lstParam_Modèles_Temp.AddItem lstParam_Modèles_REMDOC.List(iRow)
        End If
    Next iRow
    Set rs = Nothing
    Set oDoc = Nothing
    oWord.Quit
    Set oWord = Nothing
    btnControle.Enabled = True
    btnControle.BackColor = &HC0E0FF
    Me.MousePointer = vbDefault
    MsgBox "Fin du contrôle..."
    lstParam_Modèles_Temp.Clear
    lstParam_Modèles_Temp.ForeColor = ColorConstants.vbBlack
    lstParam_Modèles_Temp.FontBold = False
    Call lstParam_Modèles_Init
    
End Sub

Private Sub btnImprimer_Click()
Dim X As Long
Dim aFont As String
Dim aSize As Long

    On Error GoTo errHandler
    If lstParam_BIATABK2.ListCount > 0 Then
        CmDialog.PrinterDefault = True
        CmDialog.CancelError = True
        CmDialog.flags = cdlPDReturnDC + cdlPDNoPageNums + cdlPDDisablePrintToFile
        CmDialog.ShowPrinter
        aFont = Printer.FontName
        aSize = Printer.FontSize
        Printer.FontName = "Courier New"
        Printer.FontSize = 12
        For X = 0 To lstParam_BIATABK2.ListCount - 1
            If InStr(LCase(lstParam_BIATABK2.List(X)), "ajouter un enregistrement") <= 0 Then
                Printer.Print lstParam_BIATABK2.List(X)
            End If
        Next X
        Printer.EndDoc
        Printer.FontName = aFont
        Printer.FontSize = aSize
        MsgBox "Fin de l'impression..."
        Exit Sub
    End If
errHandler:
If Err = 32755 Then
    MsgBox "Impression annulée !"
Else
    MsgBox "Impression impossible !"
End If

End Sub

Private Sub cboParam_BIATABK2_Click()
txtParam_BIATABK2 = cboParam_BIATABK2
End Sub


Private Sub cmdContext_Click()
Unload Me

End Sub

Private Sub cmdInfo_M_Ok_Click()
    
    fraInfo_M.Visible = False
    cmdPrint_Courrier_Word

End Sub

Private Sub cmdInfo_M_Quit_Click()
blnPrint_Courrier_Ok = False
fraInfo_M.Visible = False
cmdPrint_Courrier_Word_Quit
End Sub

Private Sub cmdParam_Courrier_Quit_Click()
fraParam_Courrier.Visible = False
End Sub

Private Sub cmdParam_Courrier_Update_Click()
Dim K As Integer, K1 As Integer, X As String
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass
fraParam_Courrier.Visible = False

newCourrier_Des = oldCourrier_Des
newCourrier_Des.BIATABTXT = ""

For K = 1 To fgParam_Courrier.Rows - 1
    fgParam_Courrier.Row = K
    fgParam_Courrier.Col = 0
    K1 = Val(fgParam_Courrier.Text)
    fgParam_Courrier.Col = 2
    Select Case Trim(fgParam_Courrier)
        Case "Oui": Mid$(newCourrier_Des.BIATABTXT, K1, 1) = "O"
        Case "Non": Mid$(newCourrier_Des.BIATABTXT, K1, 1) = "N"
        Case Else: Mid$(newCourrier_Des.BIATABTXT, K1, 1) = " "
    End Select
Next K
Mid$(newCourrier_Des.BIATABTXT, 125, 4) = Format(Val(txtParam_Courrier_Seq), "0000")
Mid$(newCourrier_Des.BIATABTXT, 122, 1) = txtParam_Courrier_Originaux


If Not blnCourrier_Doc_Exist Then
        oldCourrier_Doc.BIATABK2 = cmdParam_Courrier_Doc_Exist(Trim(oldCourrier_Doc.BIATABTXT), True)
        blnCourrier_Doc_Exist = True
    newCourrier_Doc = oldCourrier_Doc
End If
oldCourrier_Des.BIATABK2 = oldCourrier_Doc.BIATABK2
newCourrier_Des.BIATABK2 = oldCourrier_Des.BIATABK2
If Not blnCourrier_Des_Exist Then
    Call sqlYBIATAB0_Transaction("New", newCourrier_Des, oldCourrier_Des)
    blnCourrier_Des_Exist = True
Else
    Call sqlYBIATAB0_Transaction("Update", newCourrier_Des, oldCourrier_Des)
End If
arrCourrier_Doc_Nb = 0
Me.Enabled = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
    Call MsgBox(Error, vbCritical, "cmdParam_Courrier_Update_Click")
    Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Detail_Add_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass
X = Trim(txtParam_BIATABTXT)
If X = "" Then
    Call MsgBox("Préciser le libellé", vbExclamation, "Paramétrage REMDOC")
Else
    newYBIATAB0 = oldYBIATAB0
    newYBIATAB0.BIATABK2 = Trim(txtParam_BIATABK2)
    
    If Trim(newYBIATAB0.BIATABK1) = "WINDOWS_TEMP" Then
        If Mid$(X, Len(X), 1) <> "\" Then X = X & "\"
    End If
    
    newYBIATAB0.BIATABTXT = X
    Call sqlYBIATAB0_Transaction("New", newYBIATAB0, oldYBIATAB0)
    Call lstParam_BIATABK2_Load(oldYBIATAB0.BIATABK1)
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Detail_Delete_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass
X = Trim(txtParam_BIATABK2)
If X <> Trim(oldYBIATAB0.BIATABK2) Then
    Call MsgBox("Le code a été modifié," & vbCrLf & " la suppression n'est pas possible", vbCritical, "Paramétrage")
Else

Call sqlYBIATAB0_Transaction("Delete", newYBIATAB0, oldYBIATAB0)
Call lstParam_BIATABK2_Load(oldYBIATAB0.BIATABK1)

End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Detail_Quit_Click()
fraParam_BIATABK2.Visible = False

End Sub

Private Sub cmdParam_Detail_Update_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass
X = Trim(txtParam_BIATABK2)
If X <> Trim(oldYBIATAB0.BIATABK2) Then
    Call MsgBox("Le code a été modifié," & vbCrLf & " la mise à jour n'est pas possible", vbCritical, "Paramétrage")
Else
    X = Trim(txtParam_BIATABTXT)
    If X = "" Then
        Call MsgBox("Préciser le libellé", vbExclamation, "Paramétrage REMDOC")
    Else
        newYBIATAB0 = oldYBIATAB0
    
        If Trim(newYBIATAB0.BIATABK1) = "WINDOWS_TEMP" Then
            If Mid$(X, Len(X), 1) <> "\" Then X = X & "\"
        End If
        
        newYBIATAB0.BIATABTXT = X
        
        Call sqlYBIATAB0_Transaction("Update", newYBIATAB0, oldYBIATAB0)
        Call lstParam_BIATABK2_Load(oldYBIATAB0.BIATABK1)
    End If
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdPrint_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case SSTab1.Tab
    Case 0: cmdPrint_Courrier_Init
    Case 1:
        If sstabParam.Tab = 0 Then cmdPrint_Excel
        If sstabParam.Tab = 2 Then Me.PopupMenu mnuPrint2, vbPopupMenuLeftButton

End Select
Me.Enabled = True: Me.MousePointer = 0
End Sub

Public Sub cmdPrint_Excel_YBIATAB0()
Dim xSql As String, X As String, K As Long
On Error GoTo Error_Handler


'On Error GoTo Error_Handler
'===================================================================================
With wsExcel.Cells
    .HorizontalAlignment = Excel.xlHAlignLeft
    .Font.Size = 9
    '.Font.Name = "Courier New"
End With
wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
'wsExcel.PageSetup.Zoom = 100
'wsExcel.PageSetup.Zoom = False
'wsExcel.PageSetup.FitToPagesWide = 1
'wsExcel.PageSetup.FitToPagesTall = False

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14SAB_Dossier_RDE : paramétrage" _
                                & "  (édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$D1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True
wsExcel.PageSetup.PrintArea = "$A:$D"


wsExcel.Columns(1).ColumnWidth = 10: wsExcel.Cells(1, 1) = "Id "
wsExcel.Columns(2).ColumnWidth = 15: wsExcel.Cells(1, 2) = "Nature"
wsExcel.Columns(3).ColumnWidth = 15: wsExcel.Cells(1, 3) = "Code"
wsExcel.Columns(4).ColumnWidth = 75: wsExcel.Cells(1, 4) = "Libellé"

For K = 1 To 4
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = vbWhite
Next

xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'REMDOC'" _
     & " order by BIATABK1 , BIATABK2"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF

    mXls1_Row = mXls1_Row + 1
    wsExcel.Cells(mXls1_Row, 1) = rsSab("BIATABID")
    wsExcel.Cells(mXls1_Row, 2) = rsSab("BIATABK1")
    wsExcel.Cells(mXls1_Row, 3) = rsSab("BIATABK2")
    wsExcel.Cells(mXls1_Row, 4) = " " & Trim(rsSab("BIATABTXT"))
    
    rsSab.MoveNext
Loop



'======================================================================================================

Exit_sub:
'__________________________________________________________________________________


'_____________________________
Exit Sub

Error_Handler:

End Sub

Public Sub cmdPrint_Excel()
On Error GoTo Error_Handler
Dim xSql As String
Dim X As String, wFile As String, wFilex As String
Dim blnCALCS As Boolean

On Error GoTo Error_Handler
'===================================================================================
'______________________________________________'
X = paramServer("\\CDO_Archive\")
wAmjMin = DSys

blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

If X = "" Then X = "C:\Temp\"
If Mid$(X, Len(X), 1) <> "\" Then X = X & "\"

mXls1_File = mXls1_File + 1

wFile = X & Trim("SAB_Dossier_RDE " & DSYS_Time & mXls1_File & ".xlsx")
'______________________________________________
If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "SAB_Dossier_RDE : nom du fichier d'exportation", wFile)
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
    .Title = "SAB_Dossier_RDE"
    .Subject = "Paramétrage"
End With

'__________________________________________________________________________________

'appExcel.Worksheets.Add

Set wsExcel = wbExcel.Sheets(1): wsExcel.Name = "Paramétrage"

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
'wsExcel.PageSetup.Zoom = 80
wsExcel.PageSetup.Zoom = False
wsExcel.PageSetup.FitToPagesWide = 1
wsExcel.PageSetup.FitToPagesTall = False

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14SAB_Dossier_RDE, arrêté au " & dateImp10(wAmjMin) _
                                 & vbCr

wsExcel.PageSetup.CenterHorizontally = True


wsExcel.PageSetup.PrintTitleRows = "$A1:$D1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

mXls1_Row = 1

Select Case sstabParam.Tab
    Case 0:
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

Private Sub cmdPrint_Dossier_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case SSTab1.Tab
    Case 0: cmdPrint_Dossier_Init
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

If SSTab1.Tab = 0 Then
    cmdPrint.ToolTipText = "cliquer ici pour afficher les documents Word"
Else
    cmdPrint.ToolTipText = "cliquer ici pour exporter les informations dans un fichier Excel"
End If
End Sub

Private Sub fgInfo_M_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'On Error Resume Next

If y <= fgInfo_M.RowHeightMin Then
Else
    If fgInfo_M.Rows > 1 And y < fgInfo_M.Rows * fgInfo_M.CellHeight Then
        fgInfo_M.Col = 3: arrFields_BIA_Index = Val(fgInfo_M.Text)

        If arrFields_BIA_Index = mUTI_DOC_Index Then
            If Not blnUTI_DOC_Loaded Then Call fgUTI_DOC_Load
            Load frmSAB_DossierUTIDOC
            frmSAB_DossierUTIDOC.Tag = CStr(mUTI_DOC_Index) & "|" & blnUTI_DOC_Loaded & "|" & blnUTI_DOC_Ok
            frmSAB_DossierUTIDOC.Show vbModal
            'fgUTI_DOC.Visible = True
            blnUTI_DOC_Ok = True
            fgInfo_M.Col = 2
            arrFields_BIA_Value(arrFields_BIA_Index) = ""
            'fgInfo_M.CellBackColor = mColor_G0
        Else
                fgInfo_M.Col = 1
                txtInfo_M.Top = fgInfo_M.CellTop + fgInfo_M.RowHeightMin
                txtInfo_M.Height = fgInfo_M.RowHeight(fgInfo_M.Row) * 3
                txtInfo_M.Left = fgInfo_M.CellLeft + fgInfo_M.Left
                txtInfo_M.Width = fgInfo_M.CellWidth
                fgInfo_M.Col = 2
                txtInfo_M.Width = txtInfo_M.Width + fgInfo_M.CellWidth
                txtInfo_M.Text = Trim(fgInfo_M.Text)
                txtInfo_M.Visible = True
                txtInfo_M.SetFocus
        End If
    End If
End If

End Sub









Private Sub fgParam_Recap_Display_DDS()
Dim wColor As Long, X As String
Dim arrDDS_Row(128) As Long
Dim K As Long, mCols As Integer, arrDoc_Nb As Integer
Dim I As Integer

On Error GoTo Error_Handler
fgParam_Recap.Visible = False
fgParam_Recap_Reset

fgParam_Recap.Rows = 1
fgParam_Recap.Row = 0

currentAction = "fgParam_Recap_Display"
X = "select count(*)  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'REMDOC' and BIATABK1 = 'Courrier_Doc' "
 Set rsSab = cnsab.Execute(X)
 mCols = rsSab(0) + 1
 X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'REMDOC' and BIATABK1 = 'Courrier_Doc' " _
      & " order by BIATABK2 desc"
 Set rsSab = cnsab.Execute(X)
arrDoc_Nb = Val(rsSab("BIATABK2"))
ReDim arrDOC_Col(arrDoc_Nb) As Long
 X = ""
 For K = 1 To mCols
     X = X & Format(K, "### ") & "   |"
Next K
fgParam_Recap.FormatString = "Intitulé                                                                                                                          |" & X
X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'REMDOC' and BIATABK1 = 'Courrier_Doc' order by BIATABTXT"
 Set rsSab = cnsab.Execute(X)
K = 0
Do While Not rsSab.EOF
         fgParam_Recap.Rows = fgParam_Recap.Rows + 1
         fgParam_Recap.Row = fgParam_Recap.Rows - 1
         fgParam_Recap.Col = 0: fgParam_Recap.Text = Trim(rsSab("BIATABTXT"))
         fgParam_Recap.CellForeColor = vbBlue
         K = K + 1
         fgParam_Recap.Col = K
         arrDOC_Col(Val(rsSab("BIATABK2"))) = K
         fgParam_Recap.CellBackColor = mColor_G1
         fgParam_Recap.Text = "*"
    rsSab.MoveNext
Loop
X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'REMDOC' and BIATABK1 = 'Courrier_DDS' order by BIATABK2"
 Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    fgParam_Recap.Rows = fgParam_Recap.Rows + 1
    fgParam_Recap.Row = fgParam_Recap.Rows - 1
    fgParam_Recap.Col = 0: fgParam_Recap.Text = " " & Trim(rsSab("BIATABTXT"))
    arrDDS_Row(Val(rsSab("BIATABK2"))) = fgParam_Recap.Row
    rsSab.MoveNext
Loop
X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'REMDOC' and BIATABK1 = 'Courrier_Des' order by BIATABK2"
 Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    fgParam_Recap.Col = arrDOC_Col(Val(rsSab("BIATABK2")))
    X = rsSab("BIATABTXT")
    For K = 1 To 124
        If arrDDS_Row(K) > 0 Then
            fgParam_Recap.Row = arrDDS_Row(K)
            fgParam_Recap.Text = Mid$(X, K, 1)
            Select Case Mid$(X, K, 1)
                Case " "
                Case "O": fgParam_Recap.CellBackColor = mColor_G1
                Case "N": fgParam_Recap.CellBackColor = mColor_W1
            End Select
        End If
    Next K
    fgParam_Recap.Row = arrDDS_Row(125)
    fgParam_Recap.Text = Val(Mid$(X, 125, 4))
    rsSab.MoveNext
Loop
fgParam_Recap.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgParam_Recap.Row): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    'SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub fgParam_Recap_Display_SAB()
Dim wColor As Long, X As String
Dim K As Long, mCols As Integer, arrDoc_Nb As Integer, arrSAB_Nb As Integer
Dim I As Integer

On Error GoTo Error_Handler
fgParam_Recap.Visible = False
fgParam_Recap_Reset

fgParam_Recap.Rows = 1
fgParam_Recap.Row = 0
currentAction = "fgParam_Recap_Display"
X = "select count(*)  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'REMDOC' and BIATABK1 = 'Courrier_Doc' "
 Set rsSab = cnsab.Execute(X)
 mCols = rsSab(0) + 2
 X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'REMDOC' and BIATABK1 = 'Courrier_Doc' " _
      & " order by BIATABK2 desc"
 Set rsSab = cnsab.Execute(X)
arrDoc_Nb = Val(rsSab("BIATABK2"))
ReDim arrDOC_Col(arrDoc_Nb) As Long
 X = ""
 For K = 2 To mCols
    X = X & Format(K, "### ") & "   |"
Next K
 fgParam_Recap.FormatString = "Code                           |" _
     & "Intitulé                                                                                                                          |" & X
X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'REMDOC' and BIATABK1 = 'Courrier_Doc' order by BIATABTXT"
 Set rsSab = cnsab.Execute(X)
K = 1
Do While Not rsSab.EOF
         fgParam_Recap.Rows = fgParam_Recap.Rows + 1
         fgParam_Recap.Row = fgParam_Recap.Rows - 1
         fgParam_Recap.Col = 0: fgParam_Recap.Text = Val(rsSab("BIATABK2"))
         fgParam_Recap.Col = 1: fgParam_Recap.Text = Trim(rsSab("BIATABTXT"))
         fgParam_Recap.CellForeColor = vbBlue
         K = K + 1
         fgParam_Recap.Col = K
         arrDOC_Col(Val(rsSab("BIATABK2"))) = K
         fgParam_Recap.CellBackColor = mColor_G1
         fgParam_Recap.Text = "*"
    rsSab.MoveNext
Loop
'_________________________________________________________________________________________________
X = "select count(*)  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'REMDOC' and BIATABK1 = '#SAB'"
 Set rsSab = cnsab.Execute(X)
arrSAB_Nb = rsSab(0) + 1
ReDim arrSAB_Row(arrSAB_Nb) As Long, arrSAB_Id(arrSAB_Nb) As String
X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'REMDOC' and BIATABK1 = '#SAB' order by BIATABK2"
 Set rsSab = cnsab.Execute(X)
K = 0
Do While Not rsSab.EOF
    fgParam_Recap.Rows = fgParam_Recap.Rows + 1
    fgParam_Recap.Row = fgParam_Recap.Rows - 1
    fgParam_Recap.Col = 0: fgParam_Recap.Text = " " & Trim(rsSab("BIATABK2"))
    fgParam_Recap.Col = 1: fgParam_Recap.Text = " " & Trim(rsSab("BIATABTXT"))
    K = K + 1
    arrSAB_Row(K) = fgParam_Recap.Row
    arrSAB_Id(K) = Trim(rsSab("BIATABK2"))
    rsSab.MoveNext
Loop
X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'REMDOC_#SAB'  order by BIATABK1 , BIATABK2"
 Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    fgParam_Recap.Col = arrDOC_Col(Val(rsSab("BIATABK1")))
    X = Trim(rsSab("BIATABK2"))
    For K = 1 To arrSAB_Nb
        If X = arrSAB_Id(K) Then
            fgParam_Recap.Row = arrSAB_Row(K)
            fgParam_Recap.Text = "X"
            fgParam_Recap.CellBackColor = mColor_G1
        End If
    Next K
    rsSab.MoveNext
Loop
fgParam_Recap.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgParam_Recap.Row): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub








Public Sub fgInfo_M_Reset()
fgInfo_M.Clear
fgInfo_M_Sort1 = 0: fgInfo_M_Sort2 = 0
fgInfo_M_Sort1_Old = -1
fgInfo_M_RowDisplay = 0: fgInfo_M_RowClick = 0
fgInfo_M_arrIndex = fgInfo_M.Cols - 1
blnfgInfo_M_DisplayLine = False
fgInfo_M_SortAD = 6
fgInfo_M.LeftCol = fgInfo_M.FixedCols

End Sub





















Private Sub fgUTI_DOC_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'On Error Resume Next

If y <= fgUTI_DOC.RowHeightMin Then
Else
    If fgUTI_DOC.Rows > 1 Then 'And X > 3930 Then
        Select Case X
            Case Is < 3930: fgUTI_DOC.Col = 0: mUTI_DOC_Col = 0
            Case Is < 5115: fgUTI_DOC.Col = 1: mUTI_DOC_Col = 1
            Case Is < 6315: fgUTI_DOC.Col = 2: mUTI_DOC_Col = 2
            Case Else: fgUTI_DOC.Col = 3: mUTI_DOC_Col = 3
        End Select
            fgUTI_DOC.ScrollBars = flexScrollBarNone
            txtUTI_DOC_M.Top = fgUTI_DOC.CellTop + fgUTI_DOC.CellHeight  '- 100
            txtUTI_DOC_M.Height = fgUTI_DOC.RowHeight(fgUTI_DOC.Row)
            txtUTI_DOC_M.Left = fgUTI_DOC.CellLeft + fgUTI_DOC.Left
            txtUTI_DOC_M.Width = fgUTI_DOC.CellWidth
            txtUTI_DOC_M.Text = Trim(fgUTI_DOC.Text)
            
            fgUTI_DOC.Col = 0: fgUTI_DOC.CellBackColor = fgUTI_DOC.BackColorSel '&H80C0FF
            txtUTI_DOC_M.Visible = True
            txtUTI_DOC_M.SetFocus
    End If
End If

End Sub


Private Sub Form_Activate()
Set XForm = Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
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
blnControl = True

End Sub

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Return()
 '       SendKeys "{TAB}"
End Sub


Public Sub cmdContext_Quit()
lstErr.Clear: lstErr.Height = 510 '200
ProgressBar1.Visible = False
If fraParam_Courrier.Visible Then fraParam_Courrier.Visible = False: Exit Sub
If txtInfo_M.Visible Then txtInfo_M.Visible = False: Exit Sub
If fraInfo_M.Visible Then Call cmdInfo_M_Quit_Click: Exit Sub  'fraInfo_M.Visible = False: Exit Sub
If txtUTI_DOC_M.Visible Then txtUTI_DOC_M.Visible = False: Exit Sub
If fgUTI_DOC.Visible Then fgUTI_DOC.Visible = False: txtUTI_DOC_M.Visible = False: Exit Sub

If txtFg.Visible Then txtFg.Visible = False: Exit Sub

Unload Me

End Sub

Private Sub Form_Load()

frmSAB_Dossier_RDE_Show

Set XForm = Me
Me.Left = 19000 - Me.Width
KeyPreview = True

blnControl = False
'mWindowState = Me.WindowState
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

fgParam_Courrier_FormatString = fgParam_Courrier.FormatString
fgParam_Courrier.Enabled = True
fgParam_Courrier.Visible = False

fgParam_Courrier_FormatString = fgParam_Recap.FormatString
fgParam_Recap.Enabled = True
fgParam_Recap.Visible = False

SSTab1.Tab = 0
Set fraInfo_M.Container = fraDossier
fraInfo_M.Top = WebBrowser1.Top
fraInfo_M.Left = fraDossier.Width - fraInfo_M.Width - 100


fgInfo_M_FormatString = fgInfo_M.FormatString
fgInfo_M.Enabled = True
fgInfo_M.Visible = False
fgInfo_M.Left = 120
fgInfo_M.Top = 360

fraParam_Courrier.Visible = False
fraParam_Courrier.Top = libParam_Modèles_Temp_Path.Top
fraParam_Courrier.Left = lstParam_Modèles_Temp.Left

fgUTI_DOC.Visible = False: txtUTI_DOC_M.Visible = False
fgUTI_DOC.Top = fgInfo_M.Top
fgUTI_DOC.Left = fgInfo_M.Left + fgInfo_M.Width - fgUTI_DOC.Width

End Sub


Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub



Private Sub Form_Unload(Cancel As Integer)
Dim nrep As String

On Error Resume Next

appWord.Quit False
Set appWord = Nothing

    If Dir("c:\temp\ZENCCAR0*.*", vbNormal) <> "" Then
        Kill "c:\temp\ZENCCAR0*.*"
    End If
    nrep = Dir("c:\temp\ZENCCAR0*_fichiers", vbDirectory)
    Do While nrep <> ""
        Kill "c:\temp\" & nrep & "\*.*"
        RmDir "c:\temp\" & nrep
        nrep = Dir("c:\temp\ZENCCAR0*_fichiers", vbDirectory)
    Loop

End Sub







Private Sub lstCourrier_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wText As String, K1 As Integer, wCourrier_Id As Long

If lstCourrier.Visible And X > 300 Then
    If lstCourrier.ListIndex > -1 Then
        If lstCourrier.Selected(lstCourrier.ListIndex) = True Then   'And Button = vbRightButton
            wText = RTrim(lstCourrier.Text)
            wText = Mid$(wText, 3, Len(wText) - 2)
            mnuExemplaires_K = 0
            For K1 = 1 To arrCourrier_Doc_Nb
                If wText = arrCourrier_Doc(K1) Then
                    mnuExemplaires_K = K1
                    Exit For
                End If
            Next K1
            If mnuExemplaires_K > 0 Then
                Debug.Print mnuExemplaires_K
                mnuExemplaires_1.Checked = False
                mnuExemplaires_2.Checked = False
                mnuExemplaires_3.Checked = False
                Select Case arrCourrier_Originaux_Dossier_Nb(mnuExemplaires_K)
                    Case 1: mnuExemplaires_1.Checked = True
                    Case 2: mnuExemplaires_2.Checked = True
                    Case 3: mnuExemplaires_3.Checked = True
                End Select
                Me.PopupMenu mnuExemplaires ', vbPopupMenuLeftButton
            End If
        End If
    End If
    Exit Sub
End If
End Sub


Private Sub lstParam_BIATABK2_Click()
Dim blnOk As Boolean, xSql As String
Me.Enabled = False: Me.MousePointer = vbHourglass


If lstParam_BIATABK2.ListIndex <= 0 Then
    oldYBIATAB0.BIATABK2 = ""
    oldYBIATAB0.BIATABTXT = ""
    blnOk = True
Else
    oldYBIATAB0.BIATABK2 = Mid$(lstParam_BIATABK2, 1, 12)
    
    xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = '" & oldYBIATAB0.BIATABID & "'" _
         & " and BIATABK1 = '" & oldYBIATAB0.BIATABK1 & "'  and BIATABK2 = '" & oldYBIATAB0.BIATABK2 & "'"
    Set rsSab = cnsab.Execute(xSql)
    If Not rsSab.EOF Then
        oldYBIATAB0.BIATABK2 = rsSab("BIATABK2")
        oldYBIATAB0.BIATABTXT = rsSab("BIATABTXT")
        blnOk = True
    Else
        Call MsgBox("Enregistrement non trouvé dans YBIATAB0", vbCritical, "Paramétrage")
    End If
End If
If blnOk Then
    cboParam_BIATABK2.ListIndex = -1
    txtParam_BIATABK2 = Trim(oldYBIATAB0.BIATABK2)
    txtParam_BIATABTXT = Trim(oldYBIATAB0.BIATABTXT)
    fraParam_BIATABK2.Visible = True
End If

Me.Enabled = True: Me.MousePointer = 0
End Sub



Private Sub lstParam_Modèles_REMDOC_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim xSql As String
Me.Enabled = False: Me.MousePointer = vbHourglass
fraParam_Courrier.Visible = False

oldCourrier_Doc.BIATABID = "REMDOC"
oldCourrier_Doc.BIATABK1 = "Courrier_Doc"
oldCourrier_Doc.BIATABK2 = ""
oldCourrier_Doc.BIATABTXT = lstParam_Modèles_REMDOC.Text
blnCourrier_Doc_Exist = False
xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'REMDOC'" _
     & " and BIATABK1 = 'Courrier_Doc' and BIATABTXT = '" & Trim(oldCourrier_Doc.BIATABTXT) & "'"
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
    blnCourrier_Doc_Exist = True
    oldCourrier_Doc.BIATABK2 = rsSab("BIATABK2")
End If
oldCourrier_Des.BIATABID = "REMDOC"
oldCourrier_Des.BIATABK1 = "Courrier_Des"
oldCourrier_Des.BIATABK2 = oldCourrier_Doc.BIATABK2
blnCourrier_Des_Exist = False
xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'REMDOC'" _
     & " and BIATABK1 = 'Courrier_Des' and BIATABK2 = '" & oldCourrier_Doc.BIATABK2 & "'"
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
    blnCourrier_Des_Exist = True
    oldCourrier_Des.BIATABTXT = rsSab("BIATABTXT")
Else
    oldCourrier_Des.BIATABTXT = ""
End If
Me.PopupMenu mnuParam_Modèles_REMDOC
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub lstParam_Modèles_Temp_Click()
Me.PopupMenu mnuParam_Modèles_Temp

End Sub


Private Sub lstParam_BIATABK1_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
txtParam_BIATABK2.Enabled = True
cboParam_BIATABK2.Visible = False
lblParam_BIATABTXT.Visible = False

blnParam_Update = arrHab(18)
Select Case lstParam_BIATABK1
    Case "Fax": Call lstParam_BIATABK2_Load("FAX")
                txtParam_BIATABK2.Enabled = False
                cboParam_BIATABK2.Visible = True

    Case "Téléphone Négotiateur": Call lstParam_BIATABK2_Load("TEL_NEGO")
                txtParam_BIATABK2.Enabled = False
                cboParam_BIATABK2.Visible = True
    Case "Répertoire temporaire": Call lstParam_BIATABK2_Load("WINDOWS_TEMP")
                txtParam_BIATABK2.Enabled = False
                cboParam_BIATABK2.Visible = True
    Case "# champ SAB": Call lstParam_BIATABK2_Load("#SAB")
    Case "? champ BIA à saisir": Call lstParam_BIATABK2_Load("?BIA")
            lblParam_BIATABTXT.Visible = True: lblParam_BIATABTXT.ForeColor = vbMagenta
    Case "Courrier_DDS": Call lstParam_BIATABK2_Load("Courrier_DDS"): blnParam_Update = arrHab(19)
    Case "Courrier_Doc": Call lstParam_BIATABK2_Load("Courrier_Doc"): blnParam_Update = arrHab(19)
    Case "Courrier_Des": Call lstParam_BIATABK2_Load("Courrier_Des"): blnParam_Update = arrHab(19)
    Case "CommissionFR": Call lstParam_BIATABK2_Load("CommissionFR")
    Case "CommissionGB": Call lstParam_BIATABK2_Load("CommissionGB")
    
End Select
cmdParam_Detail_Add.Visible = blnParam_Update
cmdParam_Detail_Delete.Visible = blnParam_Update
cmdParam_Detail_Update.Visible = blnParam_Update

If lstParam_BIATABK2.ListCount = 1 Then lstParam_BIATABK2_Click
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub lstParam_BIATABK2_Load(lBIATABK1 As String)
Dim xSql As String, K As Integer

fraParam_BIATABK2.Visible = False

oldYBIATAB0.BIATABID = "REMDOC"
oldYBIATAB0.BIATABK1 = lBIATABK1
oldYBIATAB0.BIATABK2 = ""
oldYBIATAB0.BIATABTXT = ""

lstParam_BIATABK2.Clear
lstParam_BIATABK2.AddItem "Ajouter un enregistrement"

xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'REMDOC'" _
     & " and BIATABK1 = '" & lBIATABK1 & "'"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    lstParam_BIATABK2.AddItem rsSab("BIATABK2") & " " & rsSab("BIATABTXT")
    rsSab.MoveNext
Loop

paramREMDOC_Init

End Sub




Private Sub lstPrinters_Click()
Dim X As String, K As Integer
On Error GoTo Exit_sub

prtCollection_Index = Val(Mid$((lstPrinters.Text), 2, 2)) - 1
If prtCollection_Index >= 0 Then
    mWord_ActivePrinter = Mid$((lstPrinters.Text), 6, Len(lstPrinters.Text) - 5)
    Call lstPrinters_Load
    On Error Resume Next
    lstPrinters_LostFocus
    cmdContext.SetFocus
End If
Exit_sub:

End Sub

Private Sub lstPrinters_GotFocus()
lstPrinters.Height = 3000
lstPrinters.FontBold = False
lstPrinters.BackColor = &HC0FFFF
End Sub

Private Sub lstPrinters_LostFocus()
lstPrinters.Height = 300
lstPrinters.FontBold = True
lstPrinters.BackColor = &HE0FFE0
End Sub



Private Sub lstPrinters_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
lstPrinters.SetFocus

End Sub

Private Sub mnuExemplaires_1_Click()
If fgParam_Courrier.Visible = True Then
    fgParam_Courrier.Col = 2
    fgParam_Courrier.Text = "1"
Else
    arrCourrier_Originaux_Dossier_Nb(mnuExemplaires_K) = 1
End If
End Sub

Private Sub mnuExemplaires_2_Click()
If fgParam_Courrier.Visible = True Then
    fgParam_Courrier.Col = 2
    fgParam_Courrier.Text = "2"
Else
    arrCourrier_Originaux_Dossier_Nb(mnuExemplaires_K) = 2
End If
End Sub


Private Sub mnuExemplaires_3_Click()
If fgParam_Courrier.Visible = True Then
    fgParam_Courrier.Col = 2
    fgParam_Courrier.Text = "3"
Else
    arrCourrier_Originaux_Dossier_Nb(mnuExemplaires_K) = 3
End If
End Sub


Private Sub mnuParam_Courrier_NOK_Click()
fgParam_Courrier.Col = 2
fgParam_Courrier.Text = "Non"
fgParam_Courrier.CellBackColor = mColor_W0

End Sub

Public Sub fgParam_Courrier_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgParam_Courrier.Visible = False
mRow = fgParam_Courrier.Row

If lRow > 0 And lRow < fgParam_Courrier.Rows Then
    fgParam_Courrier.Row = lRow
    For I = 1 To fgParam_Courrier.FixedCols Step -1
        fgParam_Courrier.Col = I: fgParam_Courrier.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgParam_Courrier.Row = mRow
    If fgParam_Courrier.Row > 0 Then
        lRow = fgParam_Courrier.Row
        fgParam_Courrier.Col = fgParam_Courrier_arrIndex
        lColor_Old = fgParam_Courrier.CellBackColor
        For I = 1 To fgParam_Courrier.FixedCols Step -1
          fgParam_Courrier.Col = I: fgParam_Courrier.CellBackColor = lColor
        Next I
    End If
End If
fgParam_Courrier.LeftCol = fgParam_Courrier.FixedCols
fgParam_Courrier.Visible = True
End Sub
Public Sub fgParam_Courrier_Reset()
fgParam_Courrier.Clear
fgParam_Courrier_Sort1 = 0: fgParam_Courrier_Sort2 = 0
fgParam_Courrier_Sort1_Old = -1
fgParam_Courrier_RowDisplay = 0: fgParam_Courrier_RowClick = 0
fgParam_Courrier_arrIndex = fgParam_Courrier.Cols - 1
blnfgParam_Courrier_DisplayLine = False
fgParam_Courrier_SortAD = 6
fgParam_Courrier.LeftCol = fgParam_Courrier.FixedCols

End Sub


Public Sub fgParam_Recap_Reset()
fgParam_Recap.Clear
fgParam_Recap_Sort1 = 0: fgParam_Recap_Sort2 = 0
fgParam_Recap_Sort1_Old = -1
fgParam_Recap_RowDisplay = 0: fgParam_Recap_RowClick = 0
fgParam_Recap_arrIndex = fgParam_Recap.Cols - 1
blnfgParam_Recap_DisplayLine = False
fgParam_Recap_SortAD = 6
fgParam_Recap.LeftCol = fgParam_Recap.FixedCols

End Sub
Public Sub fgParam_Courrier_Sort()
If fgParam_Courrier.Rows > 1 Then
    fgParam_Courrier.Row = 1
    fgParam_Courrier.RowSel = fgParam_Courrier.Rows - 1
    
    If fgParam_Courrier_Sort1_Old = fgParam_Courrier_Sort1 Then
        If fgParam_Courrier_SortAD = 5 Then
            fgParam_Courrier_SortAD = 6
        Else
            fgParam_Courrier_SortAD = 5
        End If
    Else
        fgParam_Courrier_SortAD = 5
    End If
    fgParam_Courrier_Sort1_Old = fgParam_Courrier_Sort1
    
    fgParam_Courrier.Col = fgParam_Courrier_Sort1
    fgParam_Courrier.ColSel = fgParam_Courrier_Sort2
    fgParam_Courrier.Sort = fgParam_Courrier_SortAD
End If

End Sub



Public Sub fgParam_Courrier_SortX(lK As Integer)
Dim I As Integer, X As String, wIndex As Long

For I = 1 To fgParam_Courrier.Rows - 1
    fgParam_Courrier.Row = I
    fgParam_Courrier.Col = fgParam_Courrier_arrIndex
    wIndex = Val(fgParam_Courrier.Text)
    Select Case lK
    End Select
    fgParam_Courrier.Col = fgParam_Courrier_arrIndex - 1
    fgParam_Courrier.Text = X
Next I

fgParam_Courrier_Sort1 = fgParam_Courrier_arrIndex - 1: fgParam_Courrier_Sort2 = fgParam_Courrier_arrIndex - 1
fgParam_Courrier_Sort
End Sub




Private Sub mnuParam_Courrier_OK_Click()
fgParam_Courrier.Col = 2
fgParam_Courrier.Text = "Oui"
fgParam_Courrier.CellBackColor = mColor_G1


End Sub
Private Sub mnuParam_Courrier_Z_Click()
fgParam_Courrier.Col = 2
fgParam_Courrier.Text = ""
fgParam_Courrier.CellBackColor = mColor_Y0
End Sub


Private Sub fgParam_Courrier_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

If y <= fgParam_Courrier.RowHeightMin Then
Else
    If fgParam_Courrier.Rows > 1 Then
        fgParam_Courrier.Col = 1
        If Mid$(fgParam_Courrier.Text, 1, 1) = "=" Then
            mnuParam_Courrier_NOK.Enabled = False
            Me.PopupMenu mnuParam_Courrier
        Else
            If Mid$(fgParam_Courrier.Text, 1, 1) = "*" Then
                mnuParam_Courrier_NOK.Enabled = True
                Me.PopupMenu mnuParam_Courrier
            End If
        End If
        
         
    End If
End If

End Sub




Private Sub mnuParam_Modèles_REMDOC_Copier_Click()
On Error GoTo Error_Handler
    Me.Enabled = False: Me.MousePointer = vbHourglass
    msFileSystem.CopyFile paramRDE_Dossier_Path_DROPI & "Modèles\" & lstParam_Modèles_REMDOC.Text, libParam_Modèles_Temp_Path & lstParam_Modèles_REMDOC.Text
    lstParam_Modèles_Init
    
    Me.Enabled = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
    Call MsgBox(Error, vbCritical, "mnuParam_Modèles_REMDOC_Copier")
    Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuParam_Modèles_REMDOC_Delete_Click()
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass
If MsgBox("Confirmez_vous la suppression du modèle : " & vbCrLf & lstParam_Modèles_REMDOC.Text, vbYesNo, "Gestion des modèles REMDOC") = vbYes Then
    msFileSystem.DeleteFile paramRDE_Dossier_Path & "Modèles\" & lstParam_Modèles_REMDOC.Text
    lstParam_Modèles_Init
    
    If blnCourrier_Doc_Exist Then Call sqlYBIATAB0_Transaction("Delete", newCourrier_Doc, oldCourrier_Doc)
    If blnCourrier_Des_Exist Then Call sqlYBIATAB0_Transaction("Delete", newCourrier_Des, oldCourrier_Des)
    oldYBIATAB0 = newCourrier_Des
    Call sqlYBIATAB0_Transaction("Delete_#SAB", newCourrier_Doc, oldCourrier_Doc)

End If
Me.Enabled = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
    Call MsgBox(Error, vbCritical, "mnuParam_Modèles_REMDOC_Delete")
    Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuParam_Modèles_REMDOC_Des_Click()
Dim X As String, K As Integer, K1 As Integer
On Error GoTo Error_Handler

    Me.Enabled = False: Me.MousePointer = vbHourglass
    fraParam_Courrier.Visible = False
    fgParam_Courrier.Visible = False
    If fgParam_Courrier.Rows <= 2 Then
        fgParam_Courrier.Rows = 1
        X = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'REMDOC'" _
         & " and BIATABK1 = 'Courrier_DDS' order by BIATABK2"
        Set rsSab = cnsab.Execute(X)
        Do While Not rsSab.EOF
            If Val(rsSab("BIATABK2")) < 125 Then
                fgParam_Courrier.Rows = fgParam_Courrier.Rows + 1
                fgParam_Courrier.Row = fgParam_Courrier.Rows - 1
                fgParam_Courrier.Col = 0: fgParam_Courrier.Text = rsSab("BIATABK2")
                fgParam_Courrier.Col = 1: fgParam_Courrier.Text = rsSab("BIATABTXT")
            End If
            rsSab.MoveNext
        Loop
    End If
    For K = 1 To fgParam_Courrier.Rows - 1
        fgParam_Courrier.Row = K
        fgParam_Courrier.Col = 0
        K1 = Val(fgParam_Courrier.Text)
        fgParam_Courrier.Col = 2
        fgParam_Courrier.CellBackColor = mColor_Y0
        Select Case Mid$(oldCourrier_Des.BIATABTXT, K1, 1)
            Case "O": fgParam_Courrier.Text = "Oui"
            Case "N": fgParam_Courrier.Text = "Non"
            Case Else: fgParam_Courrier.Text = Mid$(oldCourrier_Des.BIATABTXT, K1, 1)
        End Select
    Next K
    txtParam_Courrier_Seq = Val(Mid$(oldCourrier_Des.BIATABTXT, 125, 4))
    txtParam_Courrier_Originaux = Mid$(oldCourrier_Des.BIATABTXT, 122, 1)
    fgParam_Courrier.Visible = True
    fraParam_Courrier.Visible = True
    Me.Enabled = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
    Call MsgBox(Error, vbCritical, "mnuParam_Modèles_REMDOC_Des_Click")
    Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuParam_Modèles_REMDOC_Rename_Click()
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim K As Integer, X As String, xFrom As String, Xto As String

    K = InStr(lstParam_Modèles_REMDOC.Text, ".")
    If K > 0 Then
        X = Mid$(lstParam_Modèles_REMDOC, 1, K - 1)
    Else
        X = lstParam_Modèles_REMDOC
    End If
    Xto = Trim(InputBox("Nouveau nom du modèle:", "Gestion des modèles", X))
    If Xto <> "" Then
        xFrom = paramRDE_Dossier_Path & "Modèles\" & lstParam_Modèles_REMDOC.Text
        msFileSystem.MoveFile xFrom, Replace(xFrom, X, Xto)
        newCourrier_Doc = oldCourrier_Doc
        newCourrier_Doc.BIATABTXT = Replace(oldCourrier_Doc.BIATABTXT, X, Xto)
        Call sqlYBIATAB0_Transaction("Update", newCourrier_Doc, oldCourrier_Doc)
        lstParam_Modèles_Init
    End If
    Me.Enabled = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
    Call MsgBox(Error, vbCritical, "mnuParam_Modèles_REMDOC_Rename")
    Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuParam_Modèles_Temp_Copier_Click()
On Error GoTo Error_Handler
Dim X As String

    Me.Enabled = False: Me.MousePointer = vbHourglass
    DoEvents
    If Dir(paramRDE_Dossier_Path & "Modèles\" & lstParam_Modèles_Temp.Text) <> "" Then Kill paramRDE_Dossier_Path & "Modèles\" & lstParam_Modèles_Temp.Text
    msFileSystem.CopyFile libParam_Modèles_Temp_Path & lstParam_Modèles_Temp.Text, paramRDE_Dossier_Path & "Modèles\" & lstParam_Modèles_Temp.Text
    X = cmdParam_Courrier_Doc_Exist(lstParam_Modèles_Temp.Text, True)
    Call cmdParam_Courrier_Doc_Fields(paramRDE_Dossier_Path & "Modèles\" & lstParam_Modèles_Temp.Text, X)
    lstParam_Modèles_Init
    DoEvents
    Me.Enabled = True: Me.MousePointer = 0
    Exit Sub

Error_Handler:
    Call MsgBox(Error, vbCritical, "mnuParam_Modèles_Temp_Copier")
    Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuParam_Modèles_Temp_Delete_Click()
On Error GoTo Error_Handler

    Me.Enabled = False: Me.MousePointer = vbHourglass
    If MsgBox("Confirmez_vous la suppression du modèle : " & vbCrLf & lstParam_Modèles_REMDOC.Text, vbYesNo, "Gestion des modèles temporaires") = vbYes Then
        msFileSystem.DeleteFile libParam_Modèles_Temp_Path & lstParam_Modèles_Temp.Text
        lstParam_Modèles_Init
    End If
    Me.Enabled = True: Me.MousePointer = 0
    Exit Sub

Error_Handler:
    Call MsgBox(Error, vbCritical, "mnuParam_Modèles_Temp_Delete")
    Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuPrint2_Excel_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String

Select Case SSTab1.Tab
    Case 1:
        If sstabParam.Tab = 2 And optParam_Recap_DDS Then
            X = "Caractérisques des courriers RDE"
            Call MSflexGrid_Excel("", "REMDOC", X, fgParam_Recap, fgParam_Recap.Cols - 1)
        End If
        If sstabParam.Tab = 2 And optParam_Recap_SAB Then
            X = "Champs #SAB / courriers RDE"
            Call MSflexGrid_Excel("", "REMDOC", X, fgParam_Recap, fgParam_Recap.Cols - 1)
        End If
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint2_Mail_Click()
Dim xObjet As String, xMesg As String
Me.Enabled = False: Me.MousePointer = vbHourglass

Select Case SSTab1.Tab
    Case 1:
        If sstabParam.Tab = 2 And optParam_Recap_DDS Then
            xObjet = "Caractérisques des courriers RDE"
            xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
             & xObjet
    
            Call MSFlexGrid_SendMail(currentSSIWINMAIL, "REMDOC", xObjet, xMesg, fgParam_Recap, fgParam_Recap.Cols - 1)
        End If
        If sstabParam.Tab = 2 And optParam_Recap_SAB Then
            xObjet = "Champs #SAB / courriers RDE"
            xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
             & xObjet
    
            Call MSFlexGrid_SendMail(currentSSIWINMAIL, "REMDOC", xObjet, xMesg, fgParam_Recap, fgParam_Recap.Cols - 1)
        End If

End Select

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub optCourrier_All_Click()

    lstCourrier_Load

End Sub




Private Sub optCourrier_OUV_Click()
    
    lstCourrier_Load

End Sub





Private Sub optCourrier_REG_Click()

    lstCourrier_Load
    
End Sub

Private Sub optLangue_FR_Click()

    lstCourrier_Load

End Sub


Private Sub optLangue_GB_Click()
    
    lstCourrier_Load
    
End Sub


Private Sub optParam_Recap_DDS_Click()
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass

fgParam_Recap_Display_DDS

Me.Enabled = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
    Call MsgBox(Error, vbCritical, "optParam_Recap_DDS_Click")
    Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub optParam_Recap_SAB_Click()
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass

fgParam_Recap_Display_SAB

Me.Enabled = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
    Call MsgBox(Error, vbCritical, "optParam_Recap_DDS_Click")
    Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub optParam_Recap_Z_Click()
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass

fgParam_Recap.Visible = False

Me.Enabled = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
    Call MsgBox(Error, vbCritical, "optParam_Recap_Z_Click")
    Me.Enabled = True: Me.MousePointer = 0

End Sub



Private Sub sstabParam_Click(PreviousTab As Integer)
If sstabParam.Tab = 0 Then lstParam_BIATABK1_Init
If sstabParam.Tab = 1 Then lstParam_Modèles_Init

End Sub

Public Sub lstParam_Modèles_Init()
On Error Resume Next
Dim objFolder, objFiles
Dim fsoFile As File
Dim xSql As String
'___________________________________________________________________________________________________
lstParam_Modèles_REMDOC.Clear
Set objFolder = msFileSystem.GetFolder(paramRDE_Dossier_Path_DROPI & "Modèles")
Set objFiles = objFolder.Files
For Each fsoFile In objFiles
    If InStr(fsoFile.Type, "Document Microsoft Office Word") > 0 Then lstParam_Modèles_REMDOC.AddItem fsoFile.Name
Next
lstParam_Modèles_REMDOC.Visible = True
'___________________________________________________________________________________________________
oldYBIATAB0.BIATABID = "REMDOC"
oldYBIATAB0.BIATABK1 = "WINDOWS_TEMP"
oldYBIATAB0.BIATABK2 = usrName_UCase
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = '" & oldYBIATAB0.BIATABID & "' and BIATABK1 = '" & oldYBIATAB0.BIATABK1 & "'" _
     & " and BIATABK2 = '" & usrName_UCase & "'"
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
    libParam_Modèles_Temp_Path = Trim(rsSab("BIATABTXT"))
Else
    libParam_Modèles_Temp_Path = "C:\Temp\"
End If

lstParam_Modèles_Temp.Clear
Set objFolder = msFileSystem.GetFolder(libParam_Modèles_Temp_Path)
Set objFiles = objFolder.Files
For Each fsoFile In objFiles
    If InStr(fsoFile.Type, "Document Microsoft Office Word") > 0 Then lstParam_Modèles_Temp.AddItem fsoFile.Name
Next
lstParam_Modèles_Temp.Visible = True
End Sub


Public Sub lstParam_BIATABK1_Init()
If lstParam_BIATABK1.ListCount = 0 Then
    lstParam_BIATABK1.Clear
    lstParam_BIATABK1.AddItem "Fax"
    lstParam_BIATABK1.AddItem "Téléphone Négotiateur"
    lstParam_BIATABK1.AddItem "Répertoire temporaire"
    lstParam_BIATABK1.AddItem "# champ SAB"
    lstParam_BIATABK1.AddItem "? champ BIA à saisir"
    lstParam_BIATABK1.AddItem "Courrier_DDS"
    lstParam_BIATABK1.AddItem "Courrier_Doc"
    lstParam_BIATABK1.AddItem "Courrier_Des"
    
    lstParam_BIATABK2.Clear
    fraParam_BIATABK2.Visible = False
    
    cboParam_BIATABK2.Clear
    cboParam_BIATABK2.AddItem "*"
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN'" _
         & " and SSIDOMPRFX <> 'X' and SSIDOMUNIT = 'S10' order by SSIDOMUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        cboParam_BIATABK2.AddItem Trim(Mid$(rsSab("SSIDOMUIDX"), 1, 12))
        rsSab.MoveNext
    Loop
    
End If

End Sub

Private Sub txtInfo_M_Validate(Cancel As Boolean)
Dim blnNOk As Boolean, K As Integer
fgInfo_M.Col = 2
fgInfo_M.Text = txtInfo_M
arrFields_BIA_Value(arrFields_BIA_Index) = Trim(txtInfo_M)
fgInfo_M.CellBackColor = mColor_G0
txtInfo_M.Visible = False
cmdInfo_M_Ok_Visible
End Sub

Private Sub txtParam_BIATABK2_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub



Public Sub cmdPrint_Courrier_Info_M_Replace()
Dim X0 As String, xReplace As String, K As Integer, iLen As Integer, K2 As Integer
On Error GoTo Error_Handler
appWord.Selection.WholeStory

For K = 1 To fgInfo_M.Rows - 1
    fgInfo_M.Row = K
    fgInfo_M.Col = 0: X0 = Trim(fgInfo_M.Text)
    fgInfo_M.Col = 2: xReplace = Replace(Trim(fgInfo_M.Text), vbCrLf, vbCr)
    iLen = Len(xReplace)
    If iLen < 247 Then
        With appWord.Selection.Find
            .Text = X0
            .Replacement.Text = xReplace
            .Execute Replace:=wdReplaceAll
        End With
    Else
        Dim xSuite As String, X1 As String, xReplace_220 As String
        X1 = X0
        For K2 = 1 To iLen Step 220
            xSuite = X0 & "_" & K2
            If K2 + 220 < iLen Then
                xReplace_220 = Mid$(xReplace, K2, 220) & xSuite
            Else
                xReplace_220 = Mid$(xReplace, K2, iLen - K2 + 1)
            End If
            
            With appWord.Selection.Find
                .Text = X1
                .Replacement.Text = xReplace_220
                .Execute Replace:=wdReplaceAll
            End With
            X1 = xSuite
        Next K2
    End If
Next K

If blnUTI_DOC_Ok Then
    For K = 1 To arrUTI_DOC_Tbl_Nb
        Call cmdPrint_Courrier_Word_UTI_DOC(arrUTI_DOC_Tbl(K))
    Next K
End If

Exit Sub

Error_Handler:

Call MsgBox(Error, vbCritical, currentAction)


End Sub

Public Sub cmdPrint_Courrier_Word_Tables()
Dim oTbl As Table, K As Integer, X As String
On Error GoTo Error_Handler
arrUTI_DOC_Tbl_Nb = 0: arrUTI_COM_CR_Tbl_Nb = 0: arrUTI_COM_DB_Tbl_Nb = 0: arrUTI_COM_Escompte_Tbl_Nb = 0
arrUTI_BLOCAGE_Tbl_Nb = 0
For Each oTbl In appWord.ActiveDocument.Tables
    K = K + 1
    X = Trim((oTbl.Cell(1, 1).Range.Text))
    Select Case Mid$(X, 1, Len(X) - 2)
        Case "?UTI_DOC": arrUTI_DOC_Tbl_Nb = arrUTI_DOC_Tbl_Nb + 1: arrUTI_DOC_Tbl(arrUTI_DOC_Tbl_Nb) = K
        Case "#UTI_COM_CR": arrUTI_COM_CR_Tbl_Nb = arrUTI_COM_CR_Tbl_Nb + 1: arrUTI_COM_CR_Tbl(arrUTI_COM_CR_Tbl_Nb) = K
        Case "#UTI_COM_DB": arrUTI_COM_DB_Tbl_Nb = arrUTI_COM_DB_Tbl_Nb + 1: arrUTI_COM_DB_Tbl(arrUTI_COM_DB_Tbl_Nb) = K
        Case "?ESCOMPTE": arrUTI_COM_Escompte_Tbl_Nb = arrUTI_COM_Escompte_Tbl_Nb + 1: arrUTI_COM_Escompte_Tbl(arrUTI_COM_Escompte_Tbl_Nb) = K
        Case "#UTI_BLOCAGE": arrUTI_BLOCAGE_Tbl_Nb = arrUTI_BLOCAGE_Tbl_Nb + 1: arrUTI_BLOCAGE_Tbl(arrUTI_BLOCAGE_Tbl_Nb) = K
    End Select
Next oTbl

GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbInformation, "cmdPrint_Courrier_Word_Tables"
Exit_sub:
End Sub


















Private Sub txtParam_Courrier_Originaux_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
If KeyAscii < 49 Or KeyAscii > 51 Then KeyAscii = 0
End Sub

Private Sub txtParam_Courrier_Seq_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub



Public Function cmdParam_Courrier_Doc_Exist(lFileName As String, blnAdd As Boolean)
Dim xSql As String
Dim rsSab_Local As New ADODB.Recordset

cmdParam_Courrier_Doc_Exist = Null

xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'REMDOC'" _
     & " and BIATABK1 = 'Courrier_Doc' and BIATABTXT = '" & Trim(lFileName) & "'"
Set rsSab_Local = cnsab.Execute(xSql)
If Not rsSab_Local.EOF Then
    cmdParam_Courrier_Doc_Exist = rsSab_Local("BIATABK2")
Else
    If blnAdd Then
    
        Dim oldYBIATAB0_Local As typeYBIATAB0, newYBIATAB0_Local As typeYBIATAB0
        
        X = "select BIATABK2 from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'REMDOC'" _
             & " and BIATABK1 = 'Courrier_Doc' order by BIATABK2 desc"
        Set rsSab_Local = cnsab.Execute(X)
        oldYBIATAB0_Local.BIATABID = "REMDOC"
        oldYBIATAB0_Local.BIATABK1 = "Courrier_Doc"
        oldYBIATAB0_Local.BIATABTXT = lFileName
        If rsSab_Local.EOF Then
            oldYBIATAB0_Local.BIATABK2 = "000000000001"
        Else
            oldYBIATAB0_Local.BIATABK2 = Format(Val(rsSab_Local("BIATABK2") + 1), "000000000000")
        End If
        newYBIATAB0_Local = oldYBIATAB0_Local
        Call sqlYBIATAB0_Transaction("New", newYBIATAB0_Local, oldYBIATAB0_Local)
        
        cmdParam_Courrier_Doc_Exist = oldYBIATAB0_Local.BIATABK2
    End If
End If

End Function
Public Function cmdParam_Courrier_Doc_Fields(lFileName As String, lBIATABK1 As String)
Dim xSql As String, K As Integer
Dim rsSab_Local As New ADODB.Recordset
Dim newYBIATAB0_Local As typeYBIATAB0

On Error GoTo Error_Handler

Dim V
App_Debug = "cmdParam_Courrier_Doc_Fields "
cmdParam_Courrier_Doc_Fields = Null

newYBIATAB0_Local.BIATABID = "REMDOC_#SAB"
newYBIATAB0_Local.BIATABK1 = lBIATABK1
newYBIATAB0_Local.BIATABTXT = ""



'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
xSql = "Delete  from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'REMDOC_#SAB'" _
     & " and BIATABK1 = '" & lBIATABK1 & "'"
Call FEU_ROUGE
Set rsSab_Local = cnsab.Execute(xSql)
Call FEU_VERT
ReDim arrDOC(2), arrDOC_REF(2)
arrDOC(1) = lFileName
arrDoc_Nb = 1

Set appWord = New Word.Application

docWord_Concatenate appWord, arrDOC(), arrDoc_Nb, arrDOC_REF()
appWord.Selection.WholeStory

For K = 1 To arrFields_SAB_Nb
        With appWord.Selection.Find
            .Wrap = wdFindContinue
            .Text = arrFields_SAB_Name(K)
            .Execute
        End With
    If appWord.Selection.Find.Found Then
        newYBIATAB0_Local.BIATABK2 = arrFields_SAB_Name(K)
        V = sqlYBIATAB0_Insert(newYBIATAB0_Local)
        If Not IsNull(V) Then GoTo Error_MsgBox
    End If
Next K

'___________________________________________________________________

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    cmdParam_Courrier_Doc_Fields = V
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

     cmdPrint_Courrier_Word_Quit

End Function

























Public Sub cmdInfo_M_Ok_Visible()
Dim K As Integer, blnNOk As Boolean

For K = 1 To fgInfo_M.Rows - 1
    fgInfo_M.Row = K
    fgInfo_M.Col = 2
    If fgInfo_M.CellBackColor <> mColor_G0 Then blnNOk = True

Next K
cmdInfo_M_Ok.Visible = Not blnNOk

End Sub

Public Sub cmdPrint_Courrier_Word_SWIFT(lSWISABSWID As Long, lSWIFT_MT As String)
Dim X0 As String, wSWIFT_Text As String, X As String, xValue As String, xField As String, K As Integer, K2 As Integer, iLen As Integer, iAsc13  As Integer
Dim arrWord_Field(100) As String, arrWord_Field_Nb As Integer
On Error GoTo Error_Handler

X0 = lSWIFT_MT

X = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABSWID = " & lSWISABSWID
Set rsSabX = cnsab.Execute(X)

If rsSabX.EOF Then
    V = "Message swift non trouvé dans YSWISAB0 : " & lSWISABSWID
    GoTo Error_Display
End If

wSWIFT_Text = "?0?"
arrWord_Field(0) = "Message SWIFT émis par " & Trim(rsSabX("SWISABWBIC")) _
            & ", reçu le " & dateImp10_S(rsSabX("SWISABWAMJ")) & " " & timeImp8(rsSabX("SWISABWHMS")) & vbCr


Call arrMT_Fields_Load(rsSabX("SWISABWMTK"))

X = "select * from rtextField " _
    & "where Aid = " & rsSabX("SWISABWID1") _
    & " and text_s_umidl = " & rsSabX("SWISABWIDL") _
    & " and text_s_umidh  =  " & rsSabX("SWISABWIDH") _
    & " order by field_cnt"

Set rsSIDE_DB = cnSIDE_DB.Execute(X)
If Not rsSIDE_DB.EOF Then
    Do While Not rsSIDE_DB.EOF
        
        Select Case rsSIDE_DB("field_code")
            Case "45", "46", "47", "77":
                V = rsSIDE_DB("value_memo")
                If IsNull(V) Then V = rsSIDE_DB("value")
            Case Else:
                    V = rsSIDE_DB("value")
        End Select
        If IsNull(V) Then
            xValue = ""
        Else
            xValue = V
        End If
        arrWord_Field_Nb = arrWord_Field_Nb + 1
        xField = rsSIDE_DB("field_code") & rsSIDE_DB("field_option")
        arrWord_Field(arrWord_Field_Nb) = ":" & xField & ": " & arrMT_Fields_Scan(xField)
        wSWIFT_Text = wSWIFT_Text & "?" & arrWord_Field_Nb & "?" _
                    & vbCr & vbTab & Replace(xValue, vbCr, vbCr & vbTab) & vbCr
       
    
        rsSIDE_DB.MoveNext
    Loop
Else

    X = "select * from rtext " _
        & "where Aid = " & rsSabX("SWISABWID1") _
        & " and text_s_umidl = " & rsSabX("SWISABWIDL") _
        & " and text_s_umidh  =  " & rsSabX("SWISABWIDH")
    Set rsSIDE_DB = cnSIDE_DB.Execute(X)
    If Not rsSIDE_DB.EOF Then
        Call srvrText_GetBuffer_ODBC(rsSIDE_DB, xrText)
        
        xValue = xrText.text_data_block & Asc13
        iLen = Len(xValue)
        If Mid$(xValue, 1, 3) = Asc13 & Asc10 & ": " Then
            K = 3
        Else
            K = 1
        End If
        Do
            iAsc13 = InStr(K, xValue, Asc13)
            If iAsc13 > 0 Then
                X = Trim(Mid$(xValue, K, iAsc13 - K))
                If Mid$(X, 1, 1) <> ":" Then
                    wSWIFT_Text = wSWIFT_Text & vbCr & vbTab & Replace(Trim(Mid$(xValue, K, iAsc13 - K)), vbCr, vbTab)
                Else
                    K2 = InStr(2, X, ":")
                    If K2 > 0 Then
                        xField = Mid$(X, 2, K2 - 2)
                        'wSWIFT_Text = wSWIFT_Text & vbCr & Trim(Mid$(X, 2, K2 - 1)) & " " & arrMT_Fields_Scan(xField) _
                                & Trim(Mid$(X, K2 + 1, Len(X) - K2))
                     
                                
                                
                        arrWord_Field_Nb = arrWord_Field_Nb + 1
                        arrWord_Field(arrWord_Field_Nb) = ":" & xField & ": " & arrMT_Fields_Scan(xField)
                        wSWIFT_Text = wSWIFT_Text & vbCr & "?" & arrWord_Field_Nb & "?" _
                                    & vbCr & vbTab & Replace(Trim(Mid$(X, K2 + 1, Len(X) - K2)), vbCr, vbTab)

                    Else
                       wSWIFT_Text = wSWIFT_Text & vbCr & vbTab & Replace(Trim(Mid$(xValue, K, iAsc13 - K)), vbCr, vbTab)
                    End If
                End If
                
                K = iAsc13 + 2
            End If
         Loop Until iAsc13 = 0
    End If
        
End If

appWord.Selection.WholeStory
    iLen = Len(wSWIFT_Text)
    If iLen < 247 Then
        With appWord.Selection.Find
            '.Wrap = wdFindContinue
            .Text = X0
            .Replacement.Text = wSWIFT_Text
            .Execute Replace:=wdReplaceAll
        End With
    Else
        Dim xSuite As String, X1 As String, xReplace_220 As String
        X1 = X0
        For K2 = 1 To iLen Step 220
            xSuite = X0 & "_" & K2
            If K2 + 220 < iLen Then
                xReplace_220 = Mid$(wSWIFT_Text, K2, 220) & xSuite
            Else
                xReplace_220 = Mid$(wSWIFT_Text, K2, iLen - K2 + 1)
            End If
            
            With appWord.Selection.Find
                '.Wrap = wdFindContinue
                .Text = X1
                .Replacement.Text = xReplace_220
                .Execute Replace:=wdReplaceAll
            End With
            X1 = xSuite
        Next K2
    
    End If
    
appWord.Selection.WholeStory
    
    With appWord.Selection.Find
        .Wrap = wdFindContinue
        .Text = "?0?"
        .Execute
    End With
    If appWord.Selection.Find.Found Then
        With appWord.Selection.Range.Font
            .Bold = True
            .Underline = wdUnderlineSingle
            .Color = vbBlue
        End With
    End If
    
For K = 1 To arrWord_Field_Nb
    With appWord.Selection.Find
        .Wrap = wdFindContinue
        .Text = "?" & K & "?"
        .Execute
    End With
    If appWord.Selection.Find.Found Then
        With appWord.Selection.Range.Font
            .Bold = False
            '.Underline = wdUnderlineSingle
            .Italic = True
            .Color = wdColorGray75
        End With
    End If
    
Next K
    
For K = 0 To arrWord_Field_Nb
        With appWord.Selection.Find
            .Wrap = wdFindContinue
            .Text = "?" & K & "?"
            .Replacement.Text = arrWord_Field(K)
            .Execute Replace:=wdReplaceAll
        End With

Next K
    

Exit Sub

Error_Handler:
V = Error
Error_Display:
Call MsgBox(V, vbCritical, currentAction)

End Sub

Public Sub lstPrinters_Load()
On Error Resume Next
Dim mK As Integer, iLen As Integer, K As Integer, X As String
lstPrinters.Clear
iLen = 0
If Trim(mWord_ActivePrinter) = "" Then mWord_ActivePrinter = Printer.Devicename
mWord_ActivePrinter = UCase$(Trim(mWord_ActivePrinter))

For Each XPrt In Printers
    K = K + 1
    X = UCase$(Trim(XPrt.Devicename))
    If InStr(mWord_ActivePrinter, X) > 0 Then
        If Len(X) > iLen Then iLen = Len(X): mK = K
    End If
Next

K = 0
For Each XPrt In Printers
    K = K + 1
    If K = mK Then
         X = ">"
         mWord_ActivePrinter = UCase$(Trim(XPrt.Devicename))
    Else
        X = "-"
    End If
    lstPrinters.AddItem X & Format$(K, "00 - ") & UCase$(Trim(XPrt.Devicename))
Next
End Sub







Private Sub txtUTI_DOC_M_KeyPress(KeyAscii As Integer)

    If mUTI_DOC_Col = 3 Then Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtUTI_DOC_M_Validate(Cancel As Boolean)
Dim blnOk As Boolean, K As Integer
Dim retval As Integer

    retval = GetKeyState(vbKeyTab)
    If retval <= 0 Then Cancel = True
    fgUTI_DOC.Col = 0: fgUTI_DOC.CellBackColor = fgUTI_DOC.BackColor
    txtUTI_DOC_M.Visible = False
    fgUTI_DOC.Col = mUTI_DOC_Col
    If mUTI_DOC_Col = 3 Then
        fgUTI_DOC.Text = Format(Val(txtUTI_DOC_M), "0000")
    Else
        fgUTI_DOC.Text = Trim(txtUTI_DOC_M)
        If Trim(fgUTI_DOC.Text) <> "" Then
            blnOk = True
        Else
            Select Case mUTI_DOC_Col
                Case 1: fgUTI_DOC.Col = 2: If Trim(fgUTI_DOC.Text) <> "" Then blnOk = True
                Case 2: fgUTI_DOC.Col = 1: If Trim(fgUTI_DOC.Text) <> "" Then blnOk = True
            End Select
        End If
        If blnOk Then
            fgUTI_DOC.Col = 0: fgUTI_DOC.CellBackColor = mColor_G1
            fgUTI_DOC.Col = 1: fgUTI_DOC.CellBackColor = mColor_G1
            fgUTI_DOC.Col = 2: fgUTI_DOC.CellBackColor = mColor_G1
            fgUTI_DOC.Col = 3: fgUTI_DOC.CellBackColor = mColor_G1
        Else
            fgUTI_DOC.Col = 0: fgUTI_DOC.CellBackColor = fgUTI_DOC.BackColor
            fgUTI_DOC.Col = 1: fgUTI_DOC.CellBackColor = fgUTI_DOC.BackColor
            fgUTI_DOC.Col = 2: fgUTI_DOC.CellBackColor = fgUTI_DOC.BackColor
            fgUTI_DOC.Col = 3: fgUTI_DOC.CellBackColor = fgUTI_DOC.BackColor
        End If
    End If
    fgUTI_DOC.ScrollBars = flexScrollBarVertical

End Sub


