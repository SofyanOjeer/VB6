VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSAB_Dossier_CDO 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_Dossier_CDO"
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
   Icon            =   "SAB_Dossier_CDO.frx":0000
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
      ItemData        =   "SAB_Dossier_CDO.frx":030A
      Left            =   11805
      List            =   "SAB_Dossier_CDO.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   62
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
         Left            =   60
         TabIndex        =   4
         Top             =   105
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
         TabPicture(0)   =   "SAB_Dossier_CDO.frx":030E
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdPrint_Dossier"
         Tab(0).Control(1)=   "fraDossier"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Paramétrage"
         TabPicture(1)   =   "SAB_Dossier_CDO.frx":032A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sstabParam"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "."
         TabPicture(2)   =   "SAB_Dossier_CDO.frx":0346
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
            TabIndex        =   51
            Top             =   90
            Width           =   2010
         End
         Begin VB.Frame fraDossier 
            Height          =   10890
            Left            =   -74985
            TabIndex        =   24
            Top             =   420
            Width           =   15945
            Begin VB.ListBox lstMT734 
               BackColor       =   &H00E0FFFF&
               Height          =   1740
               ItemData        =   "SAB_Dossier_CDO.frx":0362
               Left            =   11535
               List            =   "SAB_Dossier_CDO.frx":0369
               Style           =   1  'Checkbox
               TabIndex        =   68
               Top             =   6570
               Visible         =   0   'False
               Width           =   4320
            End
            Begin VB.ListBox lstMT799 
               BackColor       =   &H00E0FFFF&
               Height          =   1740
               ItemData        =   "SAB_Dossier_CDO.frx":0377
               Left            =   11610
               List            =   "SAB_Dossier_CDO.frx":037E
               Style           =   1  'Checkbox
               TabIndex        =   64
               Top             =   8430
               Visible         =   0   'False
               Width           =   4320
            End
            Begin VB.ListBox lstMT707 
               BackColor       =   &H00E0FFFF&
               Height          =   1740
               ItemData        =   "SAB_Dossier_CDO.frx":038C
               Left            =   6960
               List            =   "SAB_Dossier_CDO.frx":038E
               Style           =   1  'Checkbox
               TabIndex        =   63
               Top             =   8430
               Visible         =   0   'False
               Width           =   4320
            End
            Begin MSFlexGridLib.MSFlexGrid fgZCDOMOD0 
               Height          =   7320
               Left            =   8550
               TabIndex        =   54
               Top             =   4950
               Visible         =   0   'False
               Width           =   9300
               _ExtentX        =   16404
               _ExtentY        =   12912
               _Version        =   393216
               Cols            =   10
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   16777215
               ForeColor       =   16384
               BackColorFixed  =   8421504
               ForeColorFixed  =   -2147483633
               BackColorBkg    =   -2147483633
               AllowUserResizing=   3
               FormatString    =   $"SAB_Dossier_CDO.frx":0390
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
               Height          =   3120
               ItemData        =   "SAB_Dossier_CDO.frx":044F
               Left            =   8685
               List            =   "SAB_Dossier_CDO.frx":0456
               Style           =   1  'Checkbox
               TabIndex        =   37
               Top             =   120
               Width           =   7200
            End
            Begin VB.CheckBox chkWord_Update 
               BackColor       =   &H00FFC0FF&
               Caption         =   "modifier Word"
               Height          =   420
               Left            =   6900
               TabIndex        =   35
               Top             =   2085
               Width           =   1700
            End
            Begin VB.CheckBox chkWord_Validation 
               BackColor       =   &H00C0FFC0&
               Caption         =   "afficher Word"
               Height          =   525
               Left            =   6900
               TabIndex        =   34
               Top             =   1590
               Width           =   1700
            End
            Begin VB.CheckBox chkPDF_Display 
               BackColor       =   &H00F0FFFF&
               Caption         =   "afficher         .PDF après impression"
               Height          =   690
               Left            =   6900
               TabIndex        =   33
               Top             =   2490
               Visible         =   0   'False
               Width           =   1700
            End
            Begin VB.Frame fraCourrier 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Height          =   945
               Left            =   6900
               TabIndex        =   29
               Top             =   660
               Width           =   1700
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
                  Left            =   930
                  TabIndex        =   49
                  Top             =   600
                  Width           =   700
               End
               Begin VB.OptionButton optCourrier_MOD 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "MOD"
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
                  Left            =   930
                  TabIndex        =   32
                  Top             =   180
                  Width           =   700
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
                  Left            =   135
                  TabIndex        =   31
                  Top             =   210
                  Value           =   -1  'True
                  Width           =   700
               End
               Begin VB.OptionButton optCourrier_UTI 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "UTI"
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
                  Left            =   150
                  TabIndex        =   30
                  Top             =   585
                  Width           =   700
               End
            End
            Begin VB.Frame fraLangue 
               BackColor       =   &H00F0FFFF&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   6900
               TabIndex        =   26
               Top             =   165
               Width           =   1700
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
                  Left            =   150
                  TabIndex        =   28
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
                  Left            =   975
                  TabIndex        =   27
                  Top             =   120
                  Width           =   600
               End
            End
            Begin RichTextLib.RichTextBox txtRTF 
               Height          =   10515
               Left            =   120
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   180
               Visible         =   0   'False
               Width           =   6675
               _ExtentX        =   11774
               _ExtentY        =   18547
               _Version        =   393217
               BackColor       =   15790320
               HideSelection   =   0   'False
               ReadOnly        =   -1  'True
               ScrollBars      =   3
               AutoVerbMenu    =   -1  'True
               TextRTF         =   $"SAB_Dossier_CDO.frx":0467
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
            Begin MSFlexGridLib.MSFlexGrid fgSelect 
               Height          =   7320
               Left            =   6780
               TabIndex        =   36
               Top             =   3225
               Width           =   9300
               _ExtentX        =   16404
               _ExtentY        =   12912
               _Version        =   393216
               Cols            =   6
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   16777215
               ForeColor       =   16384
               BackColorFixed  =   8421376
               ForeColorFixed  =   -2147483633
               BackColorBkg    =   -2147483633
               AllowUserResizing=   3
               FormatString    =   $"SAB_Dossier_CDO.frx":04E7
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
         Begin TabDlg.SSTab sstabParam 
            Height          =   10290
            Left            =   -74985
            TabIndex        =   8
            Top             =   435
            Width           =   15615
            _ExtentX        =   27543
            _ExtentY        =   18150
            _Version        =   393216
            Tab             =   2
            TabHeight       =   520
            TabCaption(0)   =   "Paramètres SAB_Dossier_CDO"
            TabPicture(0)   =   "SAB_Dossier_CDO.frx":05A5
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "fraParam_BIATABK2"
            Tab(0).Control(1)=   "lstParam_BIATABK2"
            Tab(0).Control(2)=   "lstParam_BIATABK1"
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Gestion des modèles Word"
            TabPicture(1)   =   "SAB_Dossier_CDO.frx":05C1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "btnControle"
            Tab(1).Control(1)=   "fraParam_Courrier"
            Tab(1).Control(2)=   "lstParam_Modèles_Temp"
            Tab(1).Control(3)=   "lstParam_Modèles_CREDOC"
            Tab(1).Control(4)=   "libParam_Modèles_CREDOC_Path"
            Tab(1).Control(5)=   "libParam_Modèles_Temp_Path"
            Tab(1).ControlCount=   6
            TabCaption(2)   =   "Tableau récapitulatif"
            TabPicture(2)   =   "SAB_Dossier_CDO.frx":05DD
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "fgParam_Recap"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "optParam_Recap_Z"
            Tab(2).Control(1).Enabled=   0   'False
            Tab(2).Control(2)=   "optParam_Recap_DDS"
            Tab(2).Control(2).Enabled=   0   'False
            Tab(2).Control(3)=   "optParam_Recap_SAB"
            Tab(2).Control(3).Enabled=   0   'False
            Tab(2).ControlCount=   4
            Begin VB.CommandButton btnControle 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Contrôle des variables non à jour"
               Height          =   555
               Left            =   -68370
               Style           =   1  'Graphical
               TabIndex        =   67
               Top             =   750
               Width           =   2025
            End
            Begin VB.OptionButton optParam_Recap_SAB 
               Caption         =   "champ #SAB / courriers"
               Height          =   255
               Left            =   6795
               TabIndex        =   58
               Top             =   525
               Width           =   2310
            End
            Begin VB.OptionButton optParam_Recap_DDS 
               Caption         =   "caractéristiques des courriers"
               Height          =   255
               Left            =   2610
               TabIndex        =   57
               Top             =   510
               Width           =   3030
            End
            Begin VB.OptionButton optParam_Recap_Z 
               Caption         =   "Aucun"
               Height          =   255
               Left            =   660
               TabIndex        =   56
               Top             =   465
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
               Left            =   -66660
               TabIndex        =   43
               Top             =   2070
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
                  Left            =   1980
                  MaxLength       =   1
                  TabIndex        =   61
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
                  TabIndex        =   47
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
                  TabIndex        =   46
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
                  TabIndex        =   45
                  Top             =   8300
                  Width           =   1170
               End
               Begin MSFlexGridLib.MSFlexGrid fgParam_Courrier 
                  Height          =   7545
                  Left            =   195
                  TabIndex        =   44
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
                  TabIndex        =   60
                  Top             =   8070
                  Width           =   1440
               End
               Begin VB.Label libParam_Courrier_Seq 
                  BackColor       =   &H0080FF80&
                  Caption         =   "Ordre d'affichage des courriers sélectionnés"
                  Height          =   465
                  Left            =   3675
                  TabIndex        =   48
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
                  Text            =   "SAB_Dossier_CDO.frx":05F9
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
                  TabIndex        =   59
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
               Left            =   -71220
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
               Left            =   -66840
               TabIndex        =   10
               Top             =   1710
               Width           =   7000
            End
            Begin VB.ListBox lstParam_Modèles_CREDOC 
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
               Left            =   -74520
               TabIndex        =   9
               Top             =   1605
               Width           =   7000
            End
            Begin MSFlexGridLib.MSFlexGrid fgParam_Recap 
               Height          =   9030
               Left            =   300
               TabIndex        =   55
               Top             =   1005
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
            Begin VB.Label libParam_Modèles_CREDOC_Path 
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
               Left            =   -74310
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
               Left            =   -66150
               TabIndex        =   11
               Top             =   800
               Width           =   5745
            End
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   10185
            Left            =   135
            TabIndex        =   5
            Top             =   405
            Width           =   15405
            _ExtentX        =   27173
            _ExtentY        =   17965
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "Tab 0"
            TabPicture(0)   =   "SAB_Dossier_CDO.frx":062E
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "txtFg"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "cboSelect_SQL"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "fraInfo_M"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Tab 1"
            TabPicture(1)   =   "SAB_Dossier_CDO.frx":064A
            Tab(1).ControlEnabled=   0   'False
            Tab(1).ControlCount=   0
            TabCaption(2)   =   "Tab 2"
            TabPicture(2)   =   "SAB_Dossier_CDO.frx":0666
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
               TabIndex        =   38
               Top             =   -420
               Width           =   11685
               Begin VB.TextBox txtUTI_Com_M 
                  BackColor       =   &H00C0E0FF&
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
                  Left            =   2415
                  MaxLength       =   2000
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   66
                  Top             =   2550
                  Visible         =   0   'False
                  Width           =   1155
               End
               Begin MSFlexGridLib.MSFlexGrid fgUTI_Com 
                  Height          =   7500
                  Left            =   3480
                  TabIndex        =   65
                  Top             =   2775
                  Visible         =   0   'False
                  Width           =   9705
                  _ExtentX        =   17119
                  _ExtentY        =   13229
                  _Version        =   393216
                  Cols            =   4
                  FixedCols       =   0
                  RowHeightMin    =   350
                  BackColor       =   16777215
                  ForeColor       =   16384
                  BackColorFixed  =   33023
                  ForeColorFixed  =   -2147483633
                  BackColorBkg    =   16448250
                  AllowUserResizing=   3
                  FormatString    =   $"SAB_Dossier_CDO.frx":0682
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
                  Height          =   315
                  Left            =   2385
                  MaxLength       =   2000
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   53
                  Top             =   1950
                  Visible         =   0   'False
                  Width           =   1155
               End
               Begin MSFlexGridLib.MSFlexGrid fgUTI_DOC 
                  Height          =   8445
                  Left            =   7800
                  TabIndex        =   52
                  Top             =   1500
                  Visible         =   0   'False
                  Width           =   7320
                  _ExtentX        =   12912
                  _ExtentY        =   14896
                  _Version        =   393216
                  Cols            =   4
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
                  TabIndex        =   42
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
                  TabIndex        =   41
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
                  TabIndex        =   39
                  Top             =   9135
                  Width           =   1605
               End
               Begin MSFlexGridLib.MSFlexGrid fgInfo_M 
                  Height          =   8520
                  Left            =   870
                  TabIndex        =   40
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
                  FormatString    =   $"SAB_Dossier_CDO.frx":0732
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
               Text            =   "SAB_Dossier_CDO.frx":0810
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
      Picture         =   "SAB_Dossier_CDO.frx":0818
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   -15
      Width           =   705
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   1245
      TabIndex        =   50
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
   Begin VB.Menu mnuParam_Modèles_CREDOC 
      Caption         =   "mnuParam_Modèles_CREDOC"
      Visible         =   0   'False
      Begin VB.Menu mnuParam_Modèles_CREDOC_Des 
         Caption         =   "Caractéristiques du courrier"
      End
      Begin VB.Menu mnuParam_Modèles_CREDOC_Copier 
         Caption         =   "Copier vers le répertoire de travail"
      End
      Begin VB.Menu mnuParam_Modèles_CREDOC_Delete 
         Caption         =   "Supprimer du répertoire de production"
      End
      Begin VB.Menu mnuParam_Modèles_CREDOC_Rename 
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
Attribute VB_Name = "frmSAB_Dossier_CDO"
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

Dim txtRTF_prtForeColor_Header As Long
Dim Mesg_aid As Long, mesg_s_umidl As Long, mesg_s_umidh As Long

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean


Dim fgZCDOMOD0_FormatString As String, fgZCDOMOD0_K As Integer
Dim fgZCDOMOD0_RowDisplay As Integer, fgZCDOMOD0_RowClick As Integer, fgZCDOMOD0_ColClick As Integer
Dim fgZCDOMOD0_ColorClick As Long, fgZCDOMOD0_ColorDisplay As Long
Dim fgZCDOMOD0_Sort1 As Integer, fgZCDOMOD0_Sort2 As Integer
Dim fgZCDOMOD0_SortAD As Integer, fgZCDOMOD0_Sort1_Old As Integer
Dim fgZCDOMOD0_arrIndex As Integer
Dim blnfgZCDOMOD0_DisplayLine As Boolean

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

Dim blnPrint_Courrier_Ok As Boolean

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
Dim mBEN_ZADRESS0 As typeZADRESS0, mBEN_TVANIFCLIT As String, mBEN_CDOTIESRN As String
Dim mBEN_Concat As String
Dim mBED_ZADRESS0 As typeZADRESS0
Dim mBED_Concat As String
Dim blnBED_ZADRESSE0 As Boolean
Dim w_ZADRESSE0 As typeZADRESS0
Dim mNOT_ZADRESS0 As typeZADRESS0
Dim mNOT_Concat As String
Dim mREM_ZADRESS0 As typeZADRESS0
Dim mREM_Concat As String

Dim mBQE_RBT_ZADRESS0 As typeZADRESS0, mBQE_RBT_Concat As String

Dim mMTD_T As String, mMTD_C As String, mMTD_D As String, mMTD_N As String
Dim mRatio_C As String, mRatio_N As String

Dim mBQE_RBT As String
Dim mSWISABSWID_707 As Long, mSWISABSWID_799 As Long, mSWISABSWID_734 As Long
Dim mSWISABSWID_707_Nb As Long, mSWISABSWID_799_Nb As Long, mSWISABSWID_734_Nb As Long


Dim mTC2_X As String, mTC2_W As String, mTC2_C As String, mTC2_N As String
Dim mECNF As typeWCDOCOM0, mENOTIF As typeWCDOCOM0
Dim blnZCDOTCO0_CDE As Boolean
Dim mELVD As typeWCDOCOM0, mIDOCIR As typeWCDOCOM0, mEPDIF As typeWCDOCOM0, mEMODIF As typeWCDOCOM0
Dim mERFA As typeWCDOCOM0, mECSIL As typeWCDOCOM0
Dim WCDOCOM0_X As typeWCDOCOM0, mCOM_OUV As String

Dim mIOUV As typeWCDOCOM0, mILVD As typeWCDOCOM0, mIMODIF As typeWCDOCOM0, mIPDIF As typeWCDOCOM0
Dim mIRFA As typeWCDOCOM0, mIACD As typeWCDOCOM0, mIACCEP As typeWCDOCOM0, mEACCEP As typeWCDOCOM0, mEACCED As typeWCDOCOM0
Dim blnZCDOTCO0_CDI As Boolean

Dim mANNEXES_NB As Long
Dim mDescription As String, mIrrégularités As String, mIrrégularités_Index As Integer
Dim mUTI_BEC As String
Dim blnZCDOUTI0_Select As Boolean
Dim mUTI_DOC_Index As Integer, blnUTI_DOC_Ok As Boolean, mUTI_DOC_Col  As Integer
Dim blnUTI_DOC_Loaded As Boolean
Dim arrUTI_DOC_Tbl_Nb As Integer, arrUTI_COM_CR_Tbl_Nb As Integer, arrUTI_COM_DB_Tbl_Nb As Integer
Dim arrUTI_COM_Escompte_Tbl_Nb As Integer, arrUTI_COM_Escompte_Tbl(100) As Integer
Dim arrUTI_DOC_Tbl(100) As Integer, arrUTI_COM_CR_Tbl(100) As Integer, arrUTI_COM_DB_Tbl(100) As Integer

Dim arrUTI_BLOCAGE_Tbl(100) As Integer, arrUTI_BLOCAGE_Tbl_Nb As Integer
Dim curUTI_BLOCAGE As Currency

Dim mUTI_Com_Index As Integer, blnUTI_Com_Ok As Boolean, mUTI_Com_Col  As Integer
Dim blnUTI_Com_Loaded As Boolean, mUTI_Com_RowMin As Integer

Dim curUTI_COM_CR As Currency, curUTI_COM_DB As Currency
Dim mREG_DVA_CR As Long, mREG_DVA_DB As Long
Dim mAR_Accord As String, mAR_Courrier As String, mATTN As String
Dim curUTI_NET_Escompte As Currency

Dim blnZCDOMOD0_Select As Boolean
Dim xZCDOMOD0 As typeZCDOMOD0
Dim mOLD_VALIDIT As String, mNEW_VALIDIT As String
Dim mOLD_CND As String, mNEW_CND As String
Dim mOLD_MTD As String, mNEW_MTD As String
Dim mOLD_MT_C As String, mNEW_MT_C As String
Dim mOLD_MT_N As String, mNEW_MT_N As String
Dim mOLD_MT_D As String, mNEW_MT_D As String
Dim mOLD_EMB As String, mNEW_EMB As String
'___________________________________________________________________________________________
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
'___________________________________________________________________________________________
Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim mXls1_Row As Long, mXls1_Cols As Long, mXls1_File As Integer
Dim mXls1_Col_1 As Long, mXls1_Col_2 As Long

Dim paramCREDOC_FAX As String, paramCREDOC_TEL_NEGO As String

Dim txtRTF_MT700 As String
Dim xrText As typerText
Dim cnSIDE_DB As New ADODB.Connection, rsSIDE_DB As New ADODB.Recordset, blnSIDE_DB_Open As Boolean

Dim mWord_ActivePrinter As String

Dim hwndWord As Long
Dim mClipBoard
Private Sub cmdPrint_Courrier_Word()
Dim X As String, K As Integer, K1 As Integer, xADR As String
Dim blnWord_Validation As Boolean, blnWord_Update As Boolean, blnPDF_Display As Boolean
Dim mPrinter_Word_Name As String, wXXX_OK As Integer, wXXX_NOK As Integer
On Error GoTo Error_Handler

currentAction = "cmdPrint_Courrier_Word"


blnWord_Update = False
blnWord_Validation = False
blnPDF_Display = False
If chkPDF_Display = "1" Then blnPDF_Display = True
If chkWord_Validation = "1" Then blnWord_Validation = True
If chkWord_Update = "1" Then blnWord_Validation = True: blnWord_Update = True


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
X = "select distinct BIATABK2 from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CREDOC_#SAB'" _
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
        
        Case "#COP_DOS":    arrFields_SAB_Value(K) = mCOP_DOS
        Case "#FAX":        arrFields_SAB_Value(K) = paramCREDOC_FAX
        Case "#JJ-MM-AAAA": arrFields_SAB_Value(K) = dateImp10_S(DSys)
        Case "#TEL_NEGO": arrFields_SAB_Value(K) = paramCREDOC_TEL_NEGO

        Case "#DEVISE": arrFields_SAB_Value(K) = xZCDODOS0.CDODOSDEV
        Case "#MONTANT": arrFields_SAB_Value(K) = Trim(Format(xZCDODOS0.CDODOSMON, "### ### ### ##0.00"))
        Case "#VALIDITE": arrFields_SAB_Value(K) = dateImp10_S(xZCDODOS0.CDODOSVAL + 19000000)
        
        Case "#BEN_CONCAT": arrFields_SAB_Value(K) = mBEN_Concat
        Case "#BEN_RS1":    arrFields_SAB_Value(K) = mBEN_ZADRESS0.ADRESSRA1
        Case "#BEN_RS2": arrFields_SAB_Value(K) = mBEN_ZADRESS0.ADRESSRA2
        Case "#BEN_ADR1": arrFields_SAB_Value(K) = mBEN_ZADRESS0.ADRESSAD1
        Case "#BEN_ADR2": arrFields_SAB_Value(K) = mBEN_ZADRESS0.ADRESSAD2
        Case "#BEN_ADR3": arrFields_SAB_Value(K) = mBEN_ZADRESS0.ADRESSAD3
        Case "#BEN_CP_VILL": arrFields_SAB_Value(K) = Trim(mBEN_ZADRESS0.ADRESSCOP) & " " & Trim(mBEN_ZADRESS0.ADRESSVIL)
        Case "#BEN_PAYS": arrFields_SAB_Value(K) = Trim(mBEN_ZADRESS0.ADRESSPAY)
        
        Case "#DON_CONCAT": arrFields_SAB_Value(K) = mDON_Concat
        Case "#DON_RS1": arrFields_SAB_Value(K) = mDON_ZADRESS0.ADRESSRA1
        Case "#DON_RS2": arrFields_SAB_Value(K) = mDON_ZADRESS0.ADRESSRA2
        Case "#DON_ADR1": arrFields_SAB_Value(K) = mDON_ZADRESS0.ADRESSAD1
        Case "#DON_ADR2": arrFields_SAB_Value(K) = mDON_ZADRESS0.ADRESSAD2
        Case "#DON_ADR3": arrFields_SAB_Value(K) = mDON_ZADRESS0.ADRESSAD3
        Case "#DON_CP_VILL": arrFields_SAB_Value(K) = Trim(mDON_ZADRESS0.ADRESSCOP) & " " & Trim(mDON_ZADRESS0.ADRESSVIL)
        Case "#DON_PAYS": arrFields_SAB_Value(K) = Trim(mDON_ZADRESS0.ADRESSPAY)
        
        Case "#BQE_REF": arrFields_SAB_Value(K) = xZCDODOS0.CDODOSEXT
        Case "#BQE_RSX": arrFields_SAB_Value(K) = Trim(mBQE_ZADRESS0.ADRESSRA1) & Trim(mBQE_ZADRESS0.ADRESSRA2)
        Case "#BQE_ZIP": arrFields_SAB_Value(K) = LTrim(Trim(mBQE_ZADRESS0.ADRESSCOP) & " " & Trim(mBQE_ZADRESS0.ADRESSVIL) & " " & Trim(mBQE_ZADRESS0.ADRESSPAY))
        Case "#BQE_ADRESSE":
                xADR = ""
                If Trim(mBQE_ZADRESS0.ADRESSRA2) <> "" Then xADR = Trim(mBQE_ZADRESS0.ADRESSRA2) & vbCr
                If Trim(mBQE_ZADRESS0.ADRESSAD1) <> "" Then xADR = xADR & Trim(mBQE_ZADRESS0.ADRESSAD1) & vbCr
                If Trim(mBQE_ZADRESS0.ADRESSAD2) <> "" Then xADR = xADR & Trim(mBQE_ZADRESS0.ADRESSAD2) & vbCr
                If Trim(mBQE_ZADRESS0.ADRESSAD3) <> "" Then xADR = xADR & Trim(mBQE_ZADRESS0.ADRESSAD3) & vbCr
                xADR = xADR & Trim(mBQE_ZADRESS0.ADRESSCOP) & " " & Trim(mBQE_ZADRESS0.ADRESSVIL)
                If Trim(mBQE_ZADRESS0.ADRESSPAY) <> "" Then xADR = xADR & vbCr & Trim(mBQE_ZADRESS0.ADRESSPAY)
                arrFields_SAB_Value(K) = xADR
        Case "#BQE_CONCAT": arrFields_SAB_Value(K) = mBQE_Concat
        Case "#BQE_RS1": arrFields_SAB_Value(K) = mBQE_ZADRESS0.ADRESSRA1
        Case "#BQE_RS2": arrFields_SAB_Value(K) = mBQE_ZADRESS0.ADRESSRA2
        Case "#BQE_ADR1": arrFields_SAB_Value(K) = mBQE_ZADRESS0.ADRESSAD1
        Case "#BQE_ADR2": arrFields_SAB_Value(K) = mBQE_ZADRESS0.ADRESSAD2
        Case "#BQE_ADR3": arrFields_SAB_Value(K) = mBQE_ZADRESS0.ADRESSAD3
        Case "#BQE_CP_VILL": arrFields_SAB_Value(K) = Trim(mBQE_ZADRESS0.ADRESSCOP) & " " & Trim(mBQE_ZADRESS0.ADRESSVIL)
        Case "#BQE_PAYS": arrFields_SAB_Value(K) = mBQE_ZADRESS0.ADRESSPAY
        Case "#BED_RS1": arrFields_SAB_Value(K) = mBED_ZADRESS0.ADRESSRA1
        Case "#BED_RS2": arrFields_SAB_Value(K) = mBED_ZADRESS0.ADRESSRA2
        Case "#BED_ADR1": arrFields_SAB_Value(K) = mBED_ZADRESS0.ADRESSAD1
        Case "#BED_ADR2": arrFields_SAB_Value(K) = mBED_ZADRESS0.ADRESSAD2
        Case "#BED_ADR3": arrFields_SAB_Value(K) = mBED_ZADRESS0.ADRESSAD3
        Case "#BED_CP_VILL": arrFields_SAB_Value(K) = Trim(mBED_ZADRESS0.ADRESSCOP) & " " & Trim(mBED_ZADRESS0.ADRESSVIL)
        Case "#BED_PAYS": arrFields_SAB_Value(K) = mBED_ZADRESS0.ADRESSPAY
         
        Case "#NOT_CONCAT": arrFields_SAB_Value(K) = mNOT_Concat
        Case "#NOT_RS1": arrFields_SAB_Value(K) = mNOT_ZADRESS0.ADRESSRA1
        Case "#NOT_RS2": arrFields_SAB_Value(K) = mNOT_ZADRESS0.ADRESSRA2
        Case "#NOT_ADR1": arrFields_SAB_Value(K) = mNOT_ZADRESS0.ADRESSAD1
        Case "#NOT_ADR2": arrFields_SAB_Value(K) = mNOT_ZADRESS0.ADRESSAD2
        Case "#NOT_ADR3": arrFields_SAB_Value(K) = mNOT_ZADRESS0.ADRESSAD3
        Case "#NOT_CP_VILL": arrFields_SAB_Value(K) = Trim(mNOT_ZADRESS0.ADRESSCOP) & " " & Trim(mNOT_ZADRESS0.ADRESSVIL)
        Case "#NOT_PAYS": arrFields_SAB_Value(K) = Trim(mNOT_ZADRESS0.ADRESSPAY)
        
        Case "#BQE_RBT_RSX": arrFields_SAB_Value(K) = Trim(mBQE_RBT_ZADRESS0.ADRESSRA1) & Trim(mBQE_RBT_ZADRESS0.ADRESSRA2)
       
        Case "#ECNF_DEV": arrFields_SAB_Value(K) = mECNF.WCDOCOMDEV
        Case "#ECNF_MTD": arrFields_SAB_Value(K) = Trim(Format(mECNF.WCDOCOMMON, "### ### ##0.00"))
        Case "#ECNF_MIN": arrFields_SAB_Value(K) = Trim(Format(mECNF.WCDOCO2MIN, "### ### ##0.00"))
        Case "#ECNF_TVA": arrFields_SAB_Value(K) = Trim(Format(mECNF.WCDOCOMMTV, "### ### ##0.00"))
        Case "#ECNF_TTC": arrFields_SAB_Value(K) = Trim(Format(mECNF.WCDOCOMMON + mECNF.WCDOCOMMTV, "### ### ##0.00"))
        Case "#ECNF_TAUX": arrFields_SAB_Value(K) = Trim(Format(mECNF.WCDOCO2TX1, "##0.00"))
        Case "#ECNF_PERX":
                Select Case mECNF.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "par trimestre indivisible"
                    Case "M": arrFields_SAB_Value(K) = "par mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
         Case "#ECNF_PER1":
                Select Case mECNF.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "pour le premier trimestre"
                    Case "M": arrFields_SAB_Value(K) = "pour le premier mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
       
        Case "#ENOTIF_DEV": arrFields_SAB_Value(K) = mENOTIF.WCDOCOMDEV
        Case "#ENOTIF_MTD": arrFields_SAB_Value(K) = Trim(Format(mENOTIF.WCDOCOMMON, "### ### ##0.00"))
        Case "#ENOTIF_MIN": arrFields_SAB_Value(K) = Trim(Format(mENOTIF.WCDOCO2MIN, "### ### ##0.00"))
        Case "#ENOTIF_TVA": arrFields_SAB_Value(K) = Trim(Format(mENOTIF.WCDOCOMMTV, "### ### ##0.00"))
        Case "#ENOTIF_TTC": arrFields_SAB_Value(K) = Trim(Format(mENOTIF.WCDOCOMMON + mENOTIF.WCDOCOMMTV, "### ### ##0.00"))
        Case "#ENOTIF_TAUX": arrFields_SAB_Value(K) = Trim(Format(mENOTIF.WCDOCO2TX1, "##0.00"))
        Case "#ENOTIF_PERX":
                Select Case mENOTIF.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "par trimestre indivisible"
                    Case "M": arrFields_SAB_Value(K) = "par mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
         Case "#ENOTIF_PER1":
                Select Case mENOTIF.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "pour le premier trimestre"
                    Case "M": arrFields_SAB_Value(K) = "pour le premier mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
        
        
        Case "#ECSIL_DEV": arrFields_SAB_Value(K) = mECSIL.WCDOCOMDEV
        Case "#ECSIL_MTD": arrFields_SAB_Value(K) = Trim(Format(mECSIL.WCDOCOMMON, "### ### ##0.00"))
        Case "#ECSIL_MIN": arrFields_SAB_Value(K) = Trim(Format(mECSIL.WCDOCO2MIN, "### ### ##0.00"))
        Case "#ECSIL_TVA": arrFields_SAB_Value(K) = Trim(Format(mECSIL.WCDOCOMMTV, "### ### ##0.00"))
        Case "#ECSIL_TTC": arrFields_SAB_Value(K) = Trim(Format(mECSIL.WCDOCOMMON + mECSIL.WCDOCOMMTV, "### ### ##0.00"))
        Case "#ECSIL_TAUX": arrFields_SAB_Value(K) = Trim(Format(mECSIL.WCDOCO2TX1, "##0.00"))
        Case "#ECSIL_PERX":
                Select Case mECSIL.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "par trimestre indivisible"
                    Case "M": arrFields_SAB_Value(K) = "par mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
         Case "#ECSIL_PER1":
                Select Case mECSIL.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "pour le premier trimestre"
                    Case "M": arrFields_SAB_Value(K) = "pour le premier mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
       
        Case "#ELVD_DEV": Call txtRTF_ZCDOTCO0_CDE: arrFields_SAB_Value(K) = mELVD.WCDOCOMDEV
        Case "#ELVD_MTD": Call txtRTF_ZCDOTCO0_CDE: arrFields_SAB_Value(K) = Trim(Format(mELVD.WCDOCOMMON, "### ### ##0.00"))
        Case "#ELVD_MIN": Call txtRTF_ZCDOTCO0_CDE: arrFields_SAB_Value(K) = Trim(Format(mELVD.WCDOCO2MIN, "### ### ##0.00"))
        Case "#ELVD_TAUX": Call txtRTF_ZCDOTCO0_CDE: arrFields_SAB_Value(K) = Trim(Format(mELVD.WCDOCO2TX1 * 10, "##0.0"))
        
        Case "#IDOCIR_DEV": Call txtRTF_ZCDOTCO0_CDE: arrFields_SAB_Value(K) = mIDOCIR.WCDOCOMDEV
        Case "#IDOCIR_MTD": Call txtRTF_ZCDOTCO0_CDE: arrFields_SAB_Value(K) = Trim(Format(mIDOCIR.WCDOCOMMON, "### ### ##0.00"))
        Case "#IDOCIR_MIN": Call txtRTF_ZCDOTCO0_CDE: arrFields_SAB_Value(K) = Trim(Format(mIDOCIR.WCDOCO2MIN, "### ### ##0.00"))
        Case "#IDOCIR_TAUX": Call txtRTF_ZCDOTCO0_CDE: arrFields_SAB_Value(K) = Trim(Format(mIDOCIR.WCDOCO2TX1, "##0.00"))
        
        Case "#EPDIF_DEV": Call txtRTF_ZCDOTCO0_CDE: arrFields_SAB_Value(K) = mEPDIF.WCDOCOMDEV
        Case "#EPDIF_MTD": Call txtRTF_ZCDOTCO0_CDE: arrFields_SAB_Value(K) = Trim(Format(mEPDIF.WCDOCOMMON, "### ### ##0.00"))
        Case "#EPDIF_TAUX":
                 Select Case mEPDIF.WCDOCO2PER
                     Case "T": arrFields_SAB_Value(K) = Trim(Format(mEPDIF.WCDOCO2TX1 / 4, "##0.00"))
                     Case "M": arrFields_SAB_Value(K) = Trim(Format(mEPDIF.WCDOCO2TX1 / 12, "##0.00"))
                     Case Else: arrFields_SAB_Value(K) = Trim(Format(mEPDIF.WCDOCO2TX1 / 10, "##0.00"))
                End Select
        Case "#EPDIF_PERX":
                Select Case mEPDIF.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "par trimestre indivisible"
                    Case "M": arrFields_SAB_Value(K) = "par mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
               Case "#EPDIF_MIN": Call txtRTF_ZCDOTCO0_CDE: arrFields_SAB_Value(K) = Trim(Format(mEPDIF.WCDOCO2MIN, "### ### ##0.00"))
        
        Case "#EMODIF_DEV": Call txtRTF_ZCDOTCO0_CDE: arrFields_SAB_Value(K) = mEMODIF.WCDOCOMDEV
        Case "#EMODIF_MTD": Call txtRTF_ZCDOTCO0_CDE: arrFields_SAB_Value(K) = Trim(Format(mEMODIF.WCDOCOMMON, "### ### ##0.00"))
        
        Case "#ERFA_DEV": Call txtRTF_ZCDOTCO0_CDE: arrFields_SAB_Value(K) = mERFA.WCDOCOMDEV
        Case "#ERFA_MTD": Call txtRTF_ZCDOTCO0_CDE: arrFields_SAB_Value(K) = Trim(Format(mERFA.WCDOCOMMON, "### ### ##0.00"))
        
        Case "#DOS_MTD_C": arrFields_SAB_Value(K) = Trim(Format(mMTD_C, "### ### ##0.00"))
        Case "#DOS_C%":  arrFields_SAB_Value(K) = mRatio_C
        Case "#DOS_MTD_N": arrFields_SAB_Value(K) = Trim(Format(mMTD_N, "### ### ##0.00"))
        Case "#DOS_N%":  arrFields_SAB_Value(K) = mRatio_N
        Case "#SWIFT_MT707":
            If mSWISABSWID_707 > 0 Then Call cmdPrint_Courrier_Word_SWIFT(mSWISABSWID_707, "#SWIFT_MT707")
        Case "#SWIFT_MT799":
            If mSWISABSWID_799 > 0 Then Call cmdPrint_Courrier_Word_SWIFT(mSWISABSWID_799, "#SWIFT_MT799")
        Case "#SWIFT_MT734":
            If mSWISABSWID_734 > 0 Then Call cmdPrint_Courrier_Word_SWIFT(mSWISABSWID_734, "#SWIFT_MT734")

'      End Select
    
'    If xZCDODOS0.CDODOSCOP = "CDI" Then
'        Select Case arrFields_SAB_Name(K)
     
        Case "#IOUV_DEV": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = mIOUV.WCDOCOMDEV
        Case "#IOUV_MTD": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIOUV.WCDOCOMMON, "### ### ##0.00"))
        Case "#IOUV_MIN": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIOUV.WCDOCO2MIN, "### ### ##0.00"))
        Case "#IOUV_TVA": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIOUV.WCDOCOMMTV, "### ### ##0.00"))
        Case "#IOUV_TTC": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIOUV.WCDOCOMMON + mIOUV.WCDOCOMMTV, "### ### ##0.00"))
        Case "#IOUV_TAUX": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIOUV.WCDOCO2TX1, "##0.00"))
        Case "#IOUV_PERX": Call txtRTF_ZCDOTCO0_CDI
                Select Case mIOUV.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "par trimestre indivisible"
                    Case "M": arrFields_SAB_Value(K) = "par mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
         Case "#IOUV_PER1": Call txtRTF_ZCDOTCO0_CDI
                Select Case mIOUV.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "pour le premier trimestre"
                    Case "M": arrFields_SAB_Value(K) = "pour le premier mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
       
       
        Case "#ILVD_DEV": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = mILVD.WCDOCOMDEV
        Case "#ILVD_MTD": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mILVD.WCDOCOMMON, "### ### ##0.00"))
        Case "#ILVD_MIN": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mILVD.WCDOCO2MIN, "### ### ##0.00"))
        Case "#ILVD_TVA": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mILVD.WCDOCOMMTV, "### ### ##0.00"))
        Case "#ILVD_TTC": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mILVD.WCDOCOMMON + mILVD.WCDOCOMMTV, "### ### ##0.00"))
        Case "#ILVD_TAUX": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mILVD.WCDOCO2TX1, "##0.00"))
        Case "#ILVD_PERX": Call txtRTF_ZCDOTCO0_CDI
                Select Case mILVD.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "par trimestre indivisible"
                    Case "M": arrFields_SAB_Value(K) = "par mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
         Case "#ILVD_PER1": Call txtRTF_ZCDOTCO0_CDI
                Select Case mILVD.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "pour le premier trimestre"
                    Case "M": arrFields_SAB_Value(K) = "pour le premier mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
       
        Case "#IRFA_DEV": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = mIRFA.WCDOCOMDEV
        Case "#IRFA_MTD": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIRFA.WCDOCOMMON, "### ### ##0.00"))
        Case "#IRFA_MIN": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIRFA.WCDOCO2MIN, "### ### ##0.00"))
        Case "#IRFA_TVA": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIRFA.WCDOCOMMTV, "### ### ##0.00"))
        Case "#IRFA_TTC": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIRFA.WCDOCOMMON + mIRFA.WCDOCOMMTV, "### ### ##0.00"))
        Case "#IRFA_TAUX": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIRFA.WCDOCO2TX1, "##0.00"))
        Case "#IRFA_PERX": Call txtRTF_ZCDOTCO0_CDI
                Select Case mIRFA.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "par trimestre indivisible"
                    Case "M": arrFields_SAB_Value(K) = "par mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
         Case "#IRFA_PER1": Call txtRTF_ZCDOTCO0_CDI
                Select Case mIRFA.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "pour le premier trimestre"
                    Case "M": arrFields_SAB_Value(K) = "pour le premier mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
       
        Case "#IMODIF_DEV": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = mIMODIF.WCDOCOMDEV
        Case "#IMODIF_MTD": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIMODIF.WCDOCOMMON, "### ### ##0.00"))
        Case "#IMODIF_MIN": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIMODIF.WCDOCO2MIN, "### ### ##0.00"))
        Case "#IMODIF_TVA": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIMODIF.WCDOCOMMTV, "### ### ##0.00"))
        Case "#IMODIF_TTC": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIMODIF.WCDOCOMMON + mIMODIF.WCDOCOMMTV, "### ### ##0.00"))
        Case "#IMODIF_TAUX": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIMODIF.WCDOCO2TX1, "##0.00"))
        Case "#IMODIF_PERX": Call txtRTF_ZCDOTCO0_CDI
                Select Case mIMODIF.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "par trimestre indivisible"
                    Case "M": arrFields_SAB_Value(K) = "par mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
         Case "#IMODIF_PER1": Call txtRTF_ZCDOTCO0_CDI
                Select Case mIMODIF.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "pour le premier trimestre"
                    Case "M": arrFields_SAB_Value(K) = "pour le premier mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
       
        Case "#IPDIF_DEV": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = mIPDIF.WCDOCOMDEV
        Case "#IPDIF_MTD": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIPDIF.WCDOCOMMON, "### ### ##0.00"))
        Case "#IPDIF_MIN": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIPDIF.WCDOCO2MIN, "### ### ##0.00"))
        Case "#IPDIF_TVA": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIPDIF.WCDOCOMMTV, "### ### ##0.00"))
        Case "#IPDIF_TTC": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIPDIF.WCDOCOMMON + mIPDIF.WCDOCOMMTV, "### ### ##0.00"))
        Case "#IPDIF_TAUX": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIPDIF.WCDOCO2TX1, "##0.00"))
        Case "#IPDIF_PERX": Call txtRTF_ZCDOTCO0_CDI
                Select Case mIPDIF.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "par trimestre indivisible"
                    Case "M": arrFields_SAB_Value(K) = "par mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
         Case "#IPDIF_PER1": Call txtRTF_ZCDOTCO0_CDI
                Select Case mIPDIF.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "pour le premier trimestre"
                    Case "M": arrFields_SAB_Value(K) = "pour le premier mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
       
        Case "#IACD_DEV": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = mIACD.WCDOCOMDEV
        Case "#IACD_MTD": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIACD.WCDOCOMMON, "### ### ##0.00"))
        Case "#IACD_MIN": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIACD.WCDOCO2MIN, "### ### ##0.00"))
        Case "#IACD_TVA": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIACD.WCDOCOMMTV, "### ### ##0.00"))
        Case "#IACD_TTC": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIACD.WCDOCOMMON + mIACD.WCDOCOMMTV, "### ### ##0.00"))
        Case "#IACD_TAUX": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIACD.WCDOCO2TX1, "##0.00"))
        Case "#IACD_PERX": Call txtRTF_ZCDOTCO0_CDI
                Select Case mIACD.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "par trimestre indivisible"
                    Case "M": arrFields_SAB_Value(K) = "par mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
         Case "#IACD_PER1": Call txtRTF_ZCDOTCO0_CDI
                Select Case mIACD.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "pour le premier trimestre"
                    Case "M": arrFields_SAB_Value(K) = "pour le premier mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
       
        Case "#IACCEP_DEV": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = mIACCEP.WCDOCOMDEV
        Case "#IACCEP_MTD": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIACCEP.WCDOCOMMON, "### ### ##0.00"))
        Case "#IACCEP_MIN": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIACCEP.WCDOCO2MIN, "### ### ##0.00"))
        Case "#IACCEP_TVA": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIACCEP.WCDOCOMMTV, "### ### ##0.00"))
        Case "#IACCEP_TTC": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIACCEP.WCDOCOMMON + mIACCEP.WCDOCOMMTV, "### ### ##0.00"))
        Case "#IACCEP_TAUX": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mIACCEP.WCDOCO2TX1, "##0.00"))
        Case "#IACCEP_PERX": Call txtRTF_ZCDOTCO0_CDI
                Select Case mIACCEP.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "par trimestre indivisible"
                    Case "M": arrFields_SAB_Value(K) = "par mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
         Case "#IACCEP_PER1": Call txtRTF_ZCDOTCO0_CDI
                Select Case mIACCEP.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "pour le premier trimestre"
                    Case "M": arrFields_SAB_Value(K) = "pour le premier mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
       
        Case "#EACCEP_DEV": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = mEACCEP.WCDOCOMDEV
        Case "#EACCEP_MTD": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mEACCEP.WCDOCOMMON, "### ### ##0.00"))
        Case "#EACCEP_MIN": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mEACCEP.WCDOCO2MIN, "### ### ##0.00"))
        Case "#EACCEP_TVA": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mEACCEP.WCDOCOMMTV, "### ### ##0.00"))
        Case "#EACCEP_TTC": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mEACCEP.WCDOCOMMON + mEACCEP.WCDOCOMMTV, "### ### ##0.00"))
        Case "#EACCEP_TAUX": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mEACCEP.WCDOCO2TX1, "##0.00"))
        Case "#EACCEP_PERX": Call txtRTF_ZCDOTCO0_CDI
                Select Case mEACCEP.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "par trimestre indivisible"
                    Case "M": arrFields_SAB_Value(K) = "par mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
         Case "#EACCEP_PER1": Call txtRTF_ZCDOTCO0_CDI
                Select Case mEACCEP.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "pour le premier trimestre"
                    Case "M": arrFields_SAB_Value(K) = "pour le premier mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
       
        Case "#EACCED_DEV": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = mEACCED.WCDOCOMDEV
        Case "#EACCED_MTD": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mEACCED.WCDOCOMMON, "### ### ##0.00"))
        Case "#EACCED_MIN": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mEACCED.WCDOCO2MIN, "### ### ##0.00"))
        Case "#EACCED_TVA": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mEACCED.WCDOCOMMTV, "### ### ##0.00"))
        Case "#EACCED_TTC": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mEACCED.WCDOCOMMON + mEACCED.WCDOCOMMTV, "### ### ##0.00"))
        Case "#EACCED_TAUX": Call txtRTF_ZCDOTCO0_CDI: arrFields_SAB_Value(K) = Trim(Format(mEACCED.WCDOCO2TX1, "##0.00"))
        Case "#EACCED_PERX": Call txtRTF_ZCDOTCO0_CDI
                Select Case mEACCED.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "par trimestre indivisible"
                    Case "M": arrFields_SAB_Value(K) = "par mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select
         Case "#EACCED_PER1": Call txtRTF_ZCDOTCO0_CDI
                Select Case mEACCED.WCDOCO2PER
                    Case "T": arrFields_SAB_Value(K) = "pour le premier trimestre"
                    Case "M": arrFields_SAB_Value(K) = "pour le premier mois"
                    Case Else: arrFields_SAB_Value(K) = ""
                End Select

'        End Select
'    End If
'    If blnZCDOUTI0_Select Then
'        Select Case arrFields_SAB_Name(K)
            
            Case "#UTI_MTD":    arrFields_SAB_Value(K) = Trim(Format(xZCDOUTI0.CDOUTIMON, "### ### ### ##0.00"))
            Case "#UTI_MDO":    arrFields_SAB_Value(K) = Trim(Format(xZCDOUTI0.CDOUTIMDO, "### ### ### ##0.00"))
            Case "#UTI_MPA":    arrFields_SAB_Value(K) = Trim(Format(xZCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"))
            Case "#UTI_NET_CR":
                Call cmdPrint_Courrier_Word_UTI_COM("C", 0, curUTI_COM_CR)
                arrFields_SAB_Value(K) = Trim(Format(xZCDOUTI0.CDOUTIMPA - curUTI_BLOCAGE - curUTI_COM_CR, "### ### ### ##0.00"))
            Case "#UTI_NET_DB":
                Call cmdPrint_Courrier_Word_UTI_COM("D", 0, curUTI_COM_DB)
                arrFields_SAB_Value(K) = Trim(Format(xZCDOUTI0.CDOUTIMPA + curUTI_COM_DB, "### ### ### ##0.00"))
            Case "#UTI_RER":    arrFields_SAB_Value(K) = Trim(xZCDOUTI0.CDOUTIRER)
            Case "#UTI_FRS":    arrFields_SAB_Value(K) = Trim(Format(xZCDOUTI0.CDOUTIMPA - xZCDOUTI0.CDOUTIMON, "### ### ### ##0.00"))
            Case "#UTI_DREM":    arrFields_SAB_Value(K) = dateImp10_S(xZCDOUTI0.CDOUTIDRE + 19000000)
            Case "#UTI_COM_CR":
                arrFields_SAB_Value(K) = "":
                For K = 1 To arrUTI_COM_CR_Tbl_Nb
                    Call cmdPrint_Courrier_Word_UTI_COM("C", arrUTI_COM_CR_Tbl(K), curUTI_COM_CR)
                Next K
            Case "#UTI_COM_DB":
                arrFields_SAB_Value(K) = "":
                 For K = 1 To arrUTI_COM_DB_Tbl_Nb
                    Call cmdPrint_Courrier_Word_UTI_COM("D", arrUTI_COM_DB_Tbl(K), curUTI_COM_DB)
                Next K
               'Call cmdPrint_Courrier_Word_UTI_COM("D", mUTI_COM_DB_Tbl, curUTI_COM_DB)
             Case "#UTI_BLOCAGE":
                arrFields_SAB_Value(K) = "":
                 For K = 1 To arrUTI_BLOCAGE_Tbl_Nb
                    Call cmdPrint_Courrier_Word_UTI_BLOCAGE("C", arrUTI_BLOCAGE_Tbl(K), curUTI_BLOCAGE)
                Next K
           Case "#REG_DVA_CR":
                    Call txtRTF_ZCDOREG0("C")
                    arrFields_SAB_Value(K) = dateImp10_S(mREG_DVA_CR + 19000000)
            Case "#REG_DVA_DB":
                    Call txtRTF_ZCDOREG0("D")
                    arrFields_SAB_Value(K) = dateImp10_S(mREG_DVA_DB + 19000000)
            Case "#UTI_NET_ESC": arrFields_SAB_Value(K) = Trim(Format(curUTI_NET_Escompte, "### ### ### ##0.00"))
        
        Case "#REM_CONCAT": arrFields_SAB_Value(K) = mREM_Concat
        Case "#REM_RS1":    arrFields_SAB_Value(K) = mREM_ZADRESS0.ADRESSRA1
        Case "#REM_RS2": arrFields_SAB_Value(K) = mREM_ZADRESS0.ADRESSRA2
        Case "#REM_ADR1": arrFields_SAB_Value(K) = mREM_ZADRESS0.ADRESSAD1
        Case "#REM_ADR2": arrFields_SAB_Value(K) = mREM_ZADRESS0.ADRESSAD2
        Case "#REM_ADR3": arrFields_SAB_Value(K) = mREM_ZADRESS0.ADRESSAD3
        Case "#REM_CP_VILL": arrFields_SAB_Value(K) = Trim(mREM_ZADRESS0.ADRESSCOP) & " " & Trim(mREM_ZADRESS0.ADRESSVIL)
        Case "#REM_PAYS": arrFields_SAB_Value(K) = Trim(mREM_ZADRESS0.ADRESSPAY)
'        End Select
'    End If
'    If blnZCDOMOD0_Select Then
'        Select Case arrFields_SAB_Name(K)
            
            Case "#OLD_MT_C":    arrFields_SAB_Value(K) = mOLD_MT_C
            Case "#OLD_MT_N":    arrFields_SAB_Value(K) = mOLD_MT_N
            Case "#OLD_MT_D":    arrFields_SAB_Value(K) = mOLD_MT_D
            Case "#OLD_MTD":    arrFields_SAB_Value(K) = mOLD_MTD
            Case "#OLD_CND":    arrFields_SAB_Value(K) = mOLD_CND
            Case "#OLD_EMB":    arrFields_SAB_Value(K) = dateImp10_S(mOLD_EMB)
            Case "#OLD_VALIDIT":    arrFields_SAB_Value(K) = dateImp10_S(mOLD_VALIDIT)
             
            Case "#NEW_MT_C":    arrFields_SAB_Value(K) = mNEW_MT_C
            Case "#NEW_MT_N":    arrFields_SAB_Value(K) = mNEW_MT_N
            Case "#NEW_MT_D":    arrFields_SAB_Value(K) = mNEW_MT_D
            Case "#NEW_MTD":    arrFields_SAB_Value(K) = mNEW_MTD
            Case "#NEW_CND":    arrFields_SAB_Value(K) = mNEW_CND
            Case "#NEW_EMB":    arrFields_SAB_Value(K) = dateImp10_S(mNEW_EMB)
            Case "#NEW_VALIDIT":    arrFields_SAB_Value(K) = dateImp10_S(mNEW_VALIDIT)
            
       'End Select
    'End If
    End Select
    rsSab.MoveNext
Loop

appWord.Selection.WholeStory

For K = 1 To arrFields_SAB_Nb
        With appWord.Selection.Find
            .Text = arrFields_SAB_Name(K)
            .Replacement.Text = arrFields_SAB_Value(K)
            .Execute Replace:=wdReplaceAll
        End With
Next K

'___________________________________________________________________
ProgressBar1.value = ProgressBar1.value + 1
With appWord.Selection.Find
    .Wrap = wdFindContinue
    .Text = "#DESCRIPTION"
    .Execute
End With
If appWord.Selection.Find.Found Then
    Dim blnDescription_255 As Boolean, xDescription_255 As String
    blnDescription_255 = False
    X = Replace(mDescription, vbCrLf, vbCr)
    Do
        If Len(X) > 255 Then
            blnDescription_255 = True
            xDescription_255 = Mid$(X, 1, 240) & "#DESCRIPTION"
            X = Mid$(X, 241, Len(X) - 240)
        Else
            blnDescription_255 = False
            xDescription_255 = X
        End If
        With appWord.Selection.Find
            
            .Text = "#DESCRIPTION"
            .Replacement.Text = xDescription_255
            .Execute Replace:=wdReplaceAll
        End With
    Loop Until Not blnDescription_255
End If
'___________________________________________________________________
ProgressBar1.value = ProgressBar1.value + 1


appWord.Selection.HomeKey Unit:=wdStory

appWord.ActivePrinter = mWord_ActivePrinter

'docWord.PrintPreview
'Call lstErr_AddItem(lstErr, cmdContext, "Filigrane ...."): DoEvents
'Call docWord_Filigrane(appWord, mCOP_DOS, WdColor.wdColorSeaGreen)
ProgressBar1.value = ProgressBar1.value + 1
If blnWord_Validation Then
    If Not blnWord_Update Then appWord.ActiveDocument.Protect Type:=wdAllowOnlyReading, NoReset:=True
    appWord.Visible = True
    appWord.Windows.Application.Activate
    appWord.Windows.Application.WindowState = wdWindowStateMaximize
    
 
    hwndWord = FindWindow(vbNullString, "Microsoft Word")
    If hwndWord <> 0 Then
        'AppActivate hWnd
        'Call SendMessage(hwnd, "blabla", 0, 0)
        'SwitchToThread
        'Me.WindowState = 1
        'SetActiveWindow hwnd
        Dim hwnd As Long
        hwnd = SetForegroundWindow(hwndWord)
    Else
       MsgBox "Impossible de trouver la fenêtre Word!", vbExclamation
    End If
    Sleep 2000
    X = MsgBox("Voulez_vous enregistrer ce document dans l'historique du courrier du dossier ?", vbYesNo, "SAB_Dossier_CDO")
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
    
    'mPrinter_Word_Name = appWord.ActivePrinter
    'appWord.ActivePrinter = Printer.Devicename
    
    Call lstErr_AddItem(lstErr, cmdContext, "Impression Client...."): DoEvents
    appWord.PrintOut
    
    Call lstErr_AddItem(lstErr, cmdContext, "Impression Dossier...."): DoEvents
    Call docWord_Filigrane(appWord, mCOP_DOS, WdColor.wdColorSeaGreen) ' .wdColorBlueGray)
    'X = "1-2"
    appWord.PrintOut , , "3", , "1", CStr(mDoc_Page_Nb)
    
    'appWord.ActivePrinter = mPrinter_Word_Name

    
    If Dir(mWord_PDF_Path) <> "" Then
        Call docWord_Filigrane(appWord, "Duplicata", WdColor.wdColorSeaGreen)
        X = xZCDODOS0.CDODOSCOP & "_" & Format(xZCDODOS0.CDODOSDOS, "000000")
        mDOS_Path = paramCDO_Dossier_Path & X
        If Not msFileSystem.FolderExists(mDOS_Path) Then MkDir mDOS_Path
        mDOS_seq = mDOS_seq + 1
        mDOS_File_pdf = mDOS_Path & "\" & X & "_" & DSYS_Time & mDOS_seq & "_" & mDOS_Modèle & ".pdf"
    
        'If Dir(mDOS_File_pdf) <> "" Then Kill mDOS_File_pdf
        Call lstErr_AddItem(lstErr, cmdContext, "Enregistrement .pdf ...."): DoEvents
    
            Call appWord.ActiveDocument.ExportAsFixedFormat(mDOS_File_pdf, wdExportFormatPDF, blnPDF_Display, wdExportOptimizeForPrint)
            
            frmSAB_Dossier.fgCourrier_Display
            frmSAB_Dossier_CDO.Hide
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

Private Sub cmdPrint_Courrier_Word_UTI_COM(lCDOREGCRD As String, lUTI_COM_Tbl As Integer, lcurTTC As Currency)
Dim K As Integer, wRows_Count As Integer, wTab_Row As Integer, WCDOCOMDEV As String
Dim xSql As String, curX As Currency, wCOM_Nb As Integer
Dim curHT_T As Currency, curTVA_T As Currency, curTTC_T As Currency
Dim curTVA As Currency
Dim wFont_Name As String, wFont_Size As Integer
Dim rsSABY As New ADODB.Recordset
Dim xCDOCOMCOM As String, xPERIODE As String
On Error GoTo Error_Handler

lcurTTC = 0




X = "select *  from " & paramIBM_Library_SAB & ".ZCDOREG0, " & paramIBM_Library_SAB & ".ZCDOCOM0" _
      & " where CDOREGETB = " & xZCDODOS0.CDODOSETB & " and CDOREGAGE = " & xZCDODOS0.CDODOSAGE _
      & " and CDOREGSER = '" & xZCDODOS0.CDODOSSER & "' and CDOREGSSE = '" & xZCDODOS0.CDODOSSSE & "'" _
      & " and CDOREGCOP = '" & xZCDODOS0.CDODOSCOP & "' and CDOREGDOS = " & xZCDODOS0.CDODOSDOS _
      & " and CDOREGNUR = " & xZCDOUTI0.CDOUTINUR & " and CDOREGUTI = " & xZCDOUTI0.CDOUTIUTI _
      & " and CDOREGCRD = '" & lCDOREGCRD & "'" _
      & " and CDOCOMETB = CDOREGETB  and CDOCOMAGE = CDOREGAGE " _
      & " and CDOCOMSER = CDOREGSER  and CDOCOMSSE = CDOREGSSE " _
      & " and CDOCOMCOP = CDOREGCOP  and CDOCOMDOS = CDOREGDOS " _
      & " and CDOCOMNUR = CDOREGNUR  and CDOCOMUTR = CDOREGUTI and CDOCOMNRE = CDOREGREG " _
      & " order by CDOCOMEVE , CDOCOMSEQ , CDOCOMSPE"
      
Set rsSabX = cnsab.Execute(X)
'________________________________________________________________________________________________
If lUTI_COM_Tbl = 0 Then
    Do Until rsSabX.EOF
        If rsSabX("CDOREGCRD") = lCDOREGCRD Then
            If rsSabX("CDOCOMMON") > 0 Then
                If lCDOREGCRD = "D" And Trim(rsSabX("CDOCOMCOM")) = "EFBEMT" Then
                    lcurTTC = lcurTTC - rsSabX("CDOCOMMON") - rsSabX("CDOCOMMTV")
                Else
                    lcurTTC = lcurTTC + rsSabX("CDOCOMMON") + rsSabX("CDOCOMMTV")
                End If
            End If
        End If
        rsSabX.MoveNext
    Loop
    
    GoTo Exit_sub
End If

'________________________________________________________________________________________________
wRows_Count = appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Rows.Count: wTab_Row = 1
wFont_Name = appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(1, 1).Range.Font.Name
wFont_Size = appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(1, 1).Range.Font.Size

Do Until rsSabX.EOF
    If rsSabX("CDOREGCRD") = lCDOREGCRD Then
        curX = rsSabX("CDOCOMMON")
        If curX > 0 Then
            curHT_T = curHT_T + curX
            wCOM_Nb = wCOM_Nb + 1
             wTab_Row = wTab_Row + 1
             If wTab_Row > wRows_Count Then
                 wRows_Count = wRows_Count + 1
                 appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Rows.Add
             End If
             xCDOCOMCOM = Trim(rsSabX("CDOCOMCOM"))
             If optLangue_FR Then
                    X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
                          & " where BIATABID = 'CREDOC' and BIATABK1 = 'CommissionFR' and BIATABK2 = '" & xCDOCOMCOM & "'"
            Else
                    X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
                          & " where BIATABID = 'CREDOC' and BIATABK1 = 'CommissionGB' and BIATABK2 = '" & xCDOCOMCOM & "'"
            End If
            Set rsSABY = cnsab.Execute(X)
            
            If xCDOCOMCOM = "ECNF" Or xCDOCOMCOM = "ECNFPT" Then
                xPERIODE = vbCrLf & " (" & dateImp10_S(rsSabX("CDOCOMDBP") + 19000000) & " - " & dateImp10_S(rsSabX("CDOCOMFNP") + 19000000) & ")"
            Else
                xPERIODE = ""
            End If
            
            If Not rsSABY.EOF Then
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Text = Trim(rsSABY("BIATABTXT")) & xPERIODE
            Else
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Text = rsSabX("CDOCOMCOM") & xPERIODE
            End If
        
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Name = wFont_Name
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Size = wFont_Size
            
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Range.Text = Format(curX, "### ### ##0.00")
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Range.Font.Name = wFont_Name
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Range.Font.Size = wFont_Size
            
            curTVA = rsSabX("CDOCOMMTV")
            If curTVA <> 0 Then
            
                X = "select *  from " & paramIBM_Library_SAB & ".ZCDOCO20" _
                    & " where CDOCO2ETB = " & xZCDODOS0.CDODOSETB & " and CDOCO2AGE = " & xZCDODOS0.CDODOSAGE _
                    & " and CDOCO2SER = '" & xZCDODOS0.CDODOSSER & "' and CDOCO2SSE = '" & xZCDODOS0.CDODOSSSE & "'" _
                    & " and CDOCO2COP = '" & xZCDODOS0.CDODOSCOP & "' and CDOCO2DOS = " & xZCDODOS0.CDODOSDOS _
                    & " and CDOCO2NUR = '" & xZCDOUTI0.CDOUTINUR & "' and CDOCO2UTI = " & xZCDOUTI0.CDOUTIUTI _
                    & " and CDOCO2EVE = '" & rsSabX("CDOCOMEVE") & "'" _
                    & " and CDOCO2SEQ = " & rsSabX("CDOCOMSEQ") _
                    & " and CDOCO2SPE = " & rsSabX("CDOCOMSPE")
                Set rsSABY = cnsab.Execute(X)
                
                If Not rsSABY.EOF Then
                    If rsSABY("CDOCO2TVA") <> "O" Then curTVA = 0
                End If
            End If
            If curTVA = 0 Then
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Text = "-"
            Else
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Text = Format(curTVA, "### ### ##0.00")
                curTVA_T = curTVA_T + curTVA
                curX = curX + curTVA
           End If
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Font.Name = wFont_Name
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Font.Size = wFont_Size
            
           
            curTTC_T = curTTC_T + curX
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Text = Format(curX, "### ### ##0.00")
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Name = wFont_Name
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Size = wFont_Size
            
            WCDOCOMDEV = rsSabX("CDOCOMDEV")
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Text = WCDOCOMDEV
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Name = wFont_Name
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Size = wFont_Size
            
            If wTab_Row Mod 2 = 0 Then
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Shading.BackgroundPatternColor = wdColorGray05
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Shading.BackgroundPatternColor = wdColorGray05
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Shading.BackgroundPatternColor = wdColorGray05
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Shading.BackgroundPatternColor = wdColorGray05
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Shading.BackgroundPatternColor = wdColorGray05
           Else
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Shading.BackgroundPatternColor = wdColorWhite
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Shading.BackgroundPatternColor = wdColorWhite
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Shading.BackgroundPatternColor = wdColorWhite
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Shading.BackgroundPatternColor = wdColorWhite
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Shading.BackgroundPatternColor = wdColorWhite
            End If
            
        End If
    End If
    rsSabX.MoveNext
Loop

If wCOM_Nb > 1 Then
    wTab_Row = wTab_Row + 1
    If wTab_Row > wRows_Count Then
        wRows_Count = wRows_Count + 1
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Rows.Add
    End If
    If optLangue_FR Then
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Text = "Montant total des frais et commissions"
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Size = wFont_Size
    Else
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Text = "Total amount of our charges and commissions"
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Size = wFont_Size - 2
    End If
    
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Name = wFont_Name
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Color = wdColorBlue
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Shading.BackgroundPatternColor = wdColorGray10
    
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Range.Text = Format(curHT_T, "### ### ##0.00")
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Range.Font.Name = wFont_Name
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Range.Font.Size = wFont_Size
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Range.Font.Color = wdColorBlue
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Shading.BackgroundPatternColor = wdColorGray05
    
    If curTVA_T = 0 Then
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Text = "-"
    Else
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Text = Format(curTVA_T, "### ### ##0.00")
    End If
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Font.Name = wFont_Name
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Font.Size = wFont_Size
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Font.Color = wdColorBlue
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Shading.BackgroundPatternColor = wdColorGray05
    
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Text = Format(curTTC_T, "### ### ##0.00")
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Name = wFont_Name
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Size = wFont_Size
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Bold = True
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Color = wdColorBlue
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Shading.BackgroundPatternColor = wdColorGray10

    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Text = WCDOCOMDEV
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Name = wFont_Name
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Size = wFont_Size
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Color = wdColorBlue
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Shading.BackgroundPatternColor = wdColorGray10
End If

lcurTTC = curTTC_T

'appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Formula Formula:="=sum(B2:B6)"
'appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Formula Formula:="=sum(C2:C6"
'appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Formula Formula:="=sum(D2:D6)"

'Dim xCur As Currency
'xCur = num_CDec(appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Text)

GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbInformation, "cmdPrint_Courrier_Word_UTI_COM"
Exit_sub:
End Sub
Private Sub cmdPrint_Courrier_Word_UTI_BLOCAGE(lCDOREGCRD As String, lUTI_COM_Tbl As Integer, lcurTTC As Currency)
Dim K As Integer, wRows_Count As Integer, wTab_Row As Integer, WCDOCOMDEV As String
Dim xSql As String, curX As Currency, wCOM_Nb As Integer
Dim curTTC_T As Currency
Dim curTVA As Currency
Dim wFont_Name As String, wFont_Size As Integer
Dim mCDOREGREG As Integer
On Error GoTo Error_Handler

lcurTTC = 0
If xZCDODOS0.CDODOSCON = "N" Then
    X = " and CDOCOMCOM = 'ENOTIF'"
Else
    X = " and CDOCOMCOM in ('ECNF' , 'ECNFPT')"
End If

X = "select *  from " & paramIBM_Library_SAB & ".ZCDOCOM0" _
      & " where CDOCOMETB = " & xZCDODOS0.CDODOSETB & " and CDOCOMAGE = " & xZCDODOS0.CDODOSAGE _
      & " and CDOCOMSER = '" & xZCDODOS0.CDODOSSER & "' and CDOCOMSSE = '" & xZCDODOS0.CDODOSSSE & "'" _
      & " and CDOCOMCOP = '" & xZCDODOS0.CDODOSCOP & "' and CDOCOMDOS = " & xZCDODOS0.CDODOSDOS _
      & X & "  and CDOCOMNRE > 0 and CDOCOMUTR = " & xZCDOUTI0.CDOUTIUTI
      
'$JPL 2015-01-26      & X & "  and CDOCOMNRE > 0"
      
Set rsSabX = cnsab.Execute(X)
If Not rsSabX.EOF Then
    mCDOREGREG = rsSabX("CDOCOMNRE")
Else
    mCDOREGREG = 2
End If



X = "select *  from " & paramIBM_Library_SAB & ".ZCDOREG0" _
      & " where CDOREGETB = " & xZCDODOS0.CDODOSETB & " and CDOREGAGE = " & xZCDODOS0.CDODOSAGE _
      & " and CDOREGSER = '" & xZCDODOS0.CDODOSSER & "' and CDOREGSSE = '" & xZCDODOS0.CDODOSSSE & "'" _
      & " and CDOREGCOP = '" & xZCDODOS0.CDODOSCOP & "' and CDOREGDOS = " & xZCDODOS0.CDODOSDOS _
      & " and CDOREGNUR = " & xZCDOUTI0.CDOUTINUR & " and CDOREGUTI = " & xZCDOUTI0.CDOUTIUTI _
      & " and CDOREGCRD = '" & lCDOREGCRD & "'" _
      & " and CDOREGREG <> " & mCDOREGREG
      
Set rsSabX = cnsab.Execute(X)
'________________________________________________________________________________________________
If lUTI_COM_Tbl = 0 Then

    GoTo Exit_sub
End If

'________________________________________________________________________________________________
wRows_Count = appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Rows.Count: wTab_Row = 1
wFont_Name = appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(1, 1).Range.Font.Name
wFont_Size = appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(1, 1).Range.Font.Size

Do Until rsSabX.EOF
    If rsSabX("CDOREGCRD") = lCDOREGCRD Then
        curX = rsSabX("CDOREGMON")
        If curX > 0 Then
            wCOM_Nb = wCOM_Nb + 1
             wTab_Row = wTab_Row + 1
             If wTab_Row > wRows_Count Then
                 wRows_Count = wRows_Count + 1
                 appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Rows.Add
             End If
        
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Text = "blocage " & cmdPrint_Courrier_Word_UTI_BLOCAGE_Intitulé
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Name = wFont_Name
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Size = wFont_Size
            
            '
            
           
            curTTC_T = curTTC_T + curX
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Text = Format(curX, "### ### ##0.00")
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Name = wFont_Name
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Size = wFont_Size
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Color = wdColorRed
            
            WCDOCOMDEV = rsSabX("CDOREGDEV")
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Text = WCDOCOMDEV
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Name = wFont_Name
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Size = wFont_Size
            
            If wTab_Row Mod 2 = 0 Then
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Shading.BackgroundPatternColor = wdColorGray05
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Shading.BackgroundPatternColor = wdColorGray05
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Shading.BackgroundPatternColor = wdColorGray05
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Shading.BackgroundPatternColor = wdColorGray05
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Shading.BackgroundPatternColor = wdColorGray05
           Else
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Shading.BackgroundPatternColor = wdColorWhite
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Shading.BackgroundPatternColor = wdColorWhite
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Shading.BackgroundPatternColor = wdColorWhite
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Shading.BackgroundPatternColor = wdColorWhite
                appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Shading.BackgroundPatternColor = wdColorWhite
            End If
            
        End If
    End If
    rsSabX.MoveNext
Loop

'If wCOM_Nb > 1 Then
    wTab_Row = wTab_Row + 1
    If wTab_Row > wRows_Count Then
        wRows_Count = wRows_Count + 1
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Rows.Add
    End If
    If optLangue_FR Then
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Text = "Montant restant à payer"
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Size = wFont_Size
    Else
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Text = "Amount to be payed"
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Size = wFont_Size - 2
    End If
    
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Name = wFont_Name
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Color = wdColorBlue
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Shading.BackgroundPatternColor = wdColorGray10
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Shading.BackgroundPatternColor = wdColorGray10
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Shading.BackgroundPatternColor = wdColorGray10
    
    
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Text = Format(xZCDOUTI0.CDOUTIMPA - curTTC_T, "### ### ##0.00")
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Name = wFont_Name
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Size = wFont_Size
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Bold = True
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Color = wdColorBlue
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Shading.BackgroundPatternColor = wdColorGray10

    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Text = WCDOCOMDEV
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Name = wFont_Name
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Size = wFont_Size
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Color = wdColorBlue
    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Shading.BackgroundPatternColor = wdColorGray10
'End If

lcurTTC = curTTC_T



GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbInformation, "cmdPrint_Courrier_Word_UTI_COM"
Exit_sub:
End Sub

Private Sub cmdPrint_Courrier_Word_UTI_COM_Escompte(lUTI_COM_Tbl As Integer)
Dim K As Integer, wRows_Count As Integer, wTab_Row As Integer, WCDOCOMDEV As String
Dim xSql As String, curX As Currency, wCOM_Nb As Integer
Dim curHT_T As Currency, curTVA_T As Currency, curTTC_T As Currency
Dim curTVA As Currency
Dim wFont_Name As String, wFont_Size As Integer

On Error GoTo Error_Handler


'________________________________________________________________________________________________
wRows_Count = appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Rows.Count: wTab_Row = 1
wFont_Name = appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(1, 1).Range.Font.Name
wFont_Size = appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(1, 1).Range.Font.Size


For K = 1 To 19
    fgUTI_Com.Row = K
    fgUTI_Com.Col = 0
    
    If fgUTI_Com.CellBackColor = 0 Or fgUTI_Com.CellBackColor = vbWhite Then
    Else
    
        wTab_Row = wTab_Row + 1
        If wTab_Row > wRows_Count Then
            wRows_Count = wRows_Count + 1
            appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Rows.Add
        End If
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Text = Trim(fgUTI_Com.Text)
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Name = wFont_Name
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Size = wFont_Size
       fgUTI_Com.Col = 1
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Range.Text = Trim(fgUTI_Com.Text)
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Range.Font.Name = wFont_Name
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Range.Font.Size = wFont_Size
        fgUTI_Com.Col = 2
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Text = Trim(fgUTI_Com.Text)
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Font.Name = wFont_Name
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Font.Size = wFont_Size
        fgUTI_Com.Col = 3
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Text = Trim(fgUTI_Com.Text)
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Name = wFont_Name
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Size = wFont_Size
               
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Text = xZCDODOS0.CDODOSDEV
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Name = wFont_Name
        appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Size = wFont_Size
        
        Select Case K
            Case 10:
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Color = wdColorBlue
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Shading.BackgroundPatternColor = wdColorGray10
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Range.Font.Color = wdColorBlue
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Shading.BackgroundPatternColor = wdColorGray10
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Font.Color = wdColorBlue
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Shading.BackgroundPatternColor = wdColorGray10
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Color = wdColorBlue
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Shading.BackgroundPatternColor = wdColorGray10
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Color = wdColorBlue
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Shading.BackgroundPatternColor = wdColorGray10
                    
                 wTab_Row = wTab_Row + 1
                If wTab_Row > wRows_Count Then
                    wRows_Count = wRows_Count + 1
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Rows.Add
                End If

                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Bold = False
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Shading.BackgroundPatternColor = &HF0FFFF
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Range.Font.Bold = False
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Shading.BackgroundPatternColor = &HF0FFFF
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Font.Bold = False
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Shading.BackgroundPatternColor = &HF0FFFF
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Bold = False
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Shading.BackgroundPatternColor = &HF0FFFF
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Bold = False
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Shading.BackgroundPatternColor = &HF0FFFF
                    
            Case 11:
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Bold = True
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Color = wdColorBlue
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Shading.BackgroundPatternColor = &HD0F0FF
                    
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Range.Font.Bold = True
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Range.Font.Color = wdColorBlue
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Shading.BackgroundPatternColor = &HD0F0FF
                    
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Font.Bold = True
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Font.Color = wdColorBlue
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Shading.BackgroundPatternColor = &HD0F0FF
                    
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Bold = True
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Color = wdColorBlue
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Shading.BackgroundPatternColor = &HD0F0FF
                    
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Bold = True
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Color = wdColorBlue
                    appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Shading.BackgroundPatternColor = &HD0F0FF
            Case Is > 11:
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Range.Font.Color = wdColorBlack
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Shading.BackgroundPatternColor = &HD0F0FF
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Range.Font.Color = wdColorRed
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Shading.BackgroundPatternColor = &HD0F0FF
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Range.Font.Color = wdColorRed
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Shading.BackgroundPatternColor = &HD0F0FF
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Range.Font.Color = wdColorRed
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Shading.BackgroundPatternColor = &HD0F0FF
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Range.Font.Color = wdColorBlack
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Shading.BackgroundPatternColor = &HD0F0FF

            Case Else:
                 If wTab_Row Mod 2 = 0 Then
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Shading.BackgroundPatternColor = wdColorGray05
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Shading.BackgroundPatternColor = wdColorGray05
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Shading.BackgroundPatternColor = wdColorGray05
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Shading.BackgroundPatternColor = wdColorGray05
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Shading.BackgroundPatternColor = wdColorGray05
                Else
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 1).Shading.BackgroundPatternColor = wdColorWhite
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 2).Shading.BackgroundPatternColor = wdColorWhite
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 3).Shading.BackgroundPatternColor = wdColorWhite
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 4).Shading.BackgroundPatternColor = wdColorWhite
                     appWord.ActiveDocument.Tables(lUTI_COM_Tbl).Cell(wTab_Row, 5).Shading.BackgroundPatternColor = wdColorWhite
                 End If
        End Select
        
    End If
Next K

GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbInformation, "cmdPrint_Courrier_Word_UTI_COM"
Exit_sub:
End Sub

Private Sub txtRTF_ZCDOREG0(lCDOREGCRD As String)
On Error GoTo Error_Handler

X = "select *  from " & paramIBM_Library_SAB & ".ZCDOREG0" _
      & " where CDOREGETB = " & xZCDODOS0.CDODOSETB & " and CDOREGAGE = " & xZCDODOS0.CDODOSAGE _
      & " and CDOREGSER = '" & xZCDODOS0.CDODOSSER & "' and CDOREGSSE = '" & xZCDODOS0.CDODOSSSE & "'" _
      & " and CDOREGCOP = '" & xZCDODOS0.CDODOSCOP & "' and CDOREGDOS = " & xZCDODOS0.CDODOSDOS _
      & " and CDOREGNUR = '" & xZCDOUTI0.CDOUTINUR & "' and CDOREGUTI = " & xZCDOUTI0.CDOUTIUTI _
      & " and CDOREGCRD = '" & lCDOREGCRD & "'"
Set rsSabX = cnsab.Execute(X)

If Not rsSabX.EOF Then
    Select Case lCDOREGCRD
        Case "C": mREG_DVA_CR = rsSabX("CDOREGDVA")
        Case "D": mREG_DVA_DB = rsSabX("CDOREGDVA")
    End Select
Else
    mREG_DVA_CR = 0
End If



GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbInformation, "cmdPrint_Courrier_Word_UTI_COM"
Exit_sub:
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
blnUTI_DOC_Ok = False: blnUTI_Com_Ok = False
curUTI_COM_CR = 0: curUTI_COM_DB = 0: curUTI_BLOCAGE = 0

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
                    arrDOC(arrDoc_Nb) = paramCDO_Dossier_Path_DROPI & "Modèles\" & X
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
    '.Wrap = wdFindContinue
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
txtRTF.SaveFile wFile

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
Dim path As String
Dim Name As String
Dim NewDoc As Boolean
Dim current As String, tmpName As String, NewDocName As String
Dim K As Integer

    NewDoc = False
    'ActiveDocument.Content.Select

    'Name = appWord.ActiveDocument.Name
    For K = 1 To lDOC_NB
                ' suppression temporaire de l'update automatique des links (évite l'apparition d'un warning message à chaque ouverture d'un fichier doc)
                ' temporary delete of links' automatic update (in order to avoid the appearance of a warning message each time a .doc file is opened)
                appWord.Options.UpdateLinksAtOpen = False
                ' Application.ScreenUpdating = False
                
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

'Extrait de l'aide à "Documents.add"
'Cet exemple montre comment créer et ouvrir un document en utilisant le modèle attaché au document actif.
                    tmpName = appWord.ActiveDocument.AttachedTemplate.FullName
                    appWord.Documents.Add Template:=tmpName, NewTemplate:=True
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
                appWord.Selection.Paste
                
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
        
    
    ' update général de toutes les données linkée
    ' general update of all linked datas
    ' Application.ScreenUpdating = True
    
    'Documents.Open FileName:="C:\Temp\vierge.docx", ConfirmConversions:=False _
    '        , ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
    '        PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
    '        WritePasswordTemplate:="", Format:=wdOpenFormatAuto, XMLTransform:=""
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
    For iRow = 0 To lstParam_Modèles_CREDOC.ListCount - 1
        oK = True
        lstParam_Modèles_CREDOC.Selected(iRow) = True
        DoEvents
        numDoc = Retourne_Num_Document("CREDOC", lstParam_Modèles_CREDOC.List(iRow))
        If numDoc <> "" Then
            Set oDoc = oWord.Documents.Open(paramCDO_Dossier_Path_DROPI & "Modèles\" & lstParam_Modèles_CREDOC.List(iRow), False, False)
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
            xSql = xSql & " where BIATABID = 'CREDOC_#SAB' and BIATABK1 = '" & numDoc & "'"
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
            If Not oK Then lstParam_Modèles_Temp.AddItem lstParam_Modèles_CREDOC.List(iRow)
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
fgUTI_DOC.Visible = False: blnUTI_DOC_Ok = False
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
        Case Else: Mid$(newCourrier_Des.BIATABTXT, K1, 1) = ""
    End Select
Next K
Mid$(newCourrier_Des.BIATABTXT, 125, 4) = Format(Val(txtParam_Courrier_Seq), "0000")
Mid$(newCourrier_Des.BIATABTXT, 122, 1) = txtParam_Courrier_Originaux


If Not blnCourrier_Doc_Exist Then
        oldCourrier_Doc.BIATABK2 = cmdParam_Courrier_Doc_Exist(Trim(oldCourrier_Doc.BIATABTXT), True)
        blnCourrier_Doc_Exist = True
    newCourrier_Doc = oldCourrier_Doc
    'Call sqlYBIATAB0_Transaction("New", newCourrier_Doc, oldCourrier_Doc)
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
    Call MsgBox("Préciser le libellé", vbExclamation, "Paramétrage CREDOC")
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
        Call MsgBox("Préciser le libellé", vbExclamation, "Paramétrage CREDOC")
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

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14SAB_Dossier_CDO : paramétrage" _
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

xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CREDOC'" _
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

wFile = X & Trim("SAB_Dossier_CDO " & DSYS_Time & mXls1_File & ".xlsx")
'______________________________________________
If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "SAB_Dossier_CDO : nom du fichier d'exportation", wFile)
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
    .Title = "SAB_Dossier_CDO"
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

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14SAB_Dossier_CDO, arrêté au " & dateImp10(wAmjMin) _
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

If fgUTI_Com.Visible Then fgUTI_Com.Visible = False: Call cmdInfo_M_Ok_Visible

If fgUTI_DOC.Visible Then Call cmdInfo_M_Ok_Visible
    fgUTI_DOC.Visible = False: txtUTI_DOC_M.Visible = False
    

If y <= fgInfo_M.RowHeightMin Then
Else
    If fgInfo_M.Rows > 1 And y < fgInfo_M.Rows * fgInfo_M.CellHeight Then
        fgInfo_M.Col = 3: arrFields_BIA_Index = Val(fgInfo_M.Text)
        
        If arrFields_BIA_Index = mUTI_DOC_Index Then
            If Not blnUTI_DOC_Loaded Then Call fgUTI_DOC_Load
            fgUTI_DOC.Visible = True
            blnUTI_DOC_Ok = True
            fgInfo_M.Col = 2
            arrFields_BIA_Value(arrFields_BIA_Index) = ""
            fgInfo_M.CellBackColor = mColor_G0
        Else
            'fgInfo_M.Col = 0
            If arrFields_BIA_Index = mUTI_Com_Index Then
                If Not blnUTI_Com_Loaded Then Call fgUTI_Com_Load
                fgInfo_M.Col = 2
                fgInfo_M.CellBackColor = mColor_G0
                blnUTI_Com_Ok = True
                fgInfo_M.Col = 1
                fgUTI_Com.Top = fgInfo_M.Top + fgInfo_M.RowHeightMin
                fgUTI_Com.Left = fgInfo_M.CellLeft + fgInfo_M.Left
                fgUTI_Com.Visible = True
                fgUTI_Com.ZOrder 0
                
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
End If

End Sub


Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgSelect.Visible = False
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
    fgSelect.Row = lRow
    For I = 2 To 0 Step -1
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        fgSelect.Col = fgSelect_arrIndex
        lColor_Old = fgSelect.CellBackColor
        For I = 2 To 0 Step -1
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
    End If
End If
fgSelect.LeftCol = fgSelect.FixedCols
fgSelect.Visible = True
End Sub


Public Sub fgZCDOMOD0_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgZCDOMOD0.Visible = False
mRow = fgZCDOMOD0.Row

If lRow > 0 And lRow < fgZCDOMOD0.Rows Then
    fgZCDOMOD0.Row = lRow
    For I = 1 To 0 Step -1
        fgZCDOMOD0.Col = I: fgZCDOMOD0.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgZCDOMOD0.Row = mRow
    If fgZCDOMOD0.Row > 0 Then
        lRow = fgZCDOMOD0.Row
        fgZCDOMOD0.Col = fgZCDOMOD0_arrIndex
        lColor_Old = fgZCDOMOD0.CellBackColor
        For I = 1 To 0 Step -1
          fgZCDOMOD0.Col = I: fgZCDOMOD0.CellBackColor = lColor
        Next I
    End If
End If
fgZCDOMOD0.LeftCol = fgZCDOMOD0.FixedCols
fgZCDOMOD0.Visible = True
End Sub



'______________________________________________________________________
Private Sub fgSelect_Display()
Dim wColor As Long, X As String

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
X = "select *  from " & paramIBM_Library_SAB & ".ZCDOUTI0 " _
      & " where CDOUTIETB = " & xZCDODOS0.CDODOSETB & " and CDOUTIAGE = " & xZCDODOS0.CDODOSAGE _
      & " and CDOUTISER = '" & xZCDODOS0.CDODOSSER & "' and CDOUTISSE = '" & xZCDODOS0.CDODOSSSE & "'" _
      & " and CDOUTICOP = '" & xZCDODOS0.CDODOSCOP & "' and CDOUTIDOS = " & xZCDODOS0.CDODOSDOS _
      & " order by CDOUTINUR , CDOUTIUTI"
      
 Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF

         fgSelect.Rows = fgSelect.Rows + 1
         fgSelect.Row = fgSelect.Rows - 1
         
         fgSelect.Col = 0: fgSelect.Text = rsSab("CDOUTIUTI")
         fgSelect.Col = 5: fgSelect.Text = rsSab("CDOUTINUR")
         fgSelect.Col = 1: fgSelect.Text = dateImp10_S(rsSab("CDOUTIPRE") + 19000000)
         fgSelect.Col = 2: fgSelect.Text = Format(rsSab("CDOUTIMON"), "### ### ### ##0.00")
         fgSelect.Col = 3:
         If rsSab("CDOUTIDCO") = "N" Then
                 Select Case Trim(rsSab("CDOUTIIRR"))
                    Case "01": fgSelect.Text = "Doc refusés"
                    Case "02": fgSelect.Text = "Demande accord"
                    Case Else: fgSelect.Text = rsSab("CDOUTIIRR")
                End Select
                fgSelect.CellBackColor = mColor_W0
        Else
            fgSelect.Text = "Oui": fgSelect.CellBackColor = mColor_G0
        End If
        
        Select Case rsSab("CDOUTIEVE")
            Case "03": X = "Utilisation"
            Case "04": X = "Acc/Ref reçu"
            Case Else: X = rsSab("CDOUTIEVE")
        End Select
        
        Select Case rsSab("CDOUTIATT")
            Case "08": X = X & "    Utilisation prête"
            Case "01": X = X & "    en attente accord/refus"
            Case "02": X = X & "    en attente levée de réserves"
            Case "09": X = X & "    totalement réglée"
            Case Else: X = X & "    ATT = " & rsSab("CDOUTIEVE")
        End Select
        
        Select Case rsSab("CDOUTIETA")
            Case "01": X = X & "    Saisie"
            Case "02": X = X & "    Validée"
            Case "03": X = X & "    Comptabilisée"
            Case "04": X = X & "    Annulation comptable"
            Case Else: X = X & "    ETA = " & rsSab("CDOUTIEVE")
        End Select
        fgSelect.Col = 4: fgSelect.Text = X
    rsSab.MoveNext
Loop

If fgSelect.Rows > 1 Then
    fgSelect.Row = fgSelect.Rows - 1
    Call fgSelect_MouseDown(1, 0, 1000, 400)
    fgSelect.TopRow = fgSelect.Row
End If

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Row): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


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
'fgParam_Recap.FormatString = fgParam_Recap_FormatString
fgParam_Recap.Row = 0

currentAction = "fgParam_Recap_Display"
X = "select count(*)  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'CREDOC' and BIATABK1 = 'Courrier_Doc' "
      
 Set rsSab = cnsab.Execute(X)
 mCols = rsSab(0) + 1
 
 X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'CREDOC' and BIATABK1 = 'Courrier_Doc' " _
      & " order by BIATABK2 desc"
      
 Set rsSab = cnsab.Execute(X)
 
arrDoc_Nb = Val(rsSab("BIATABK2"))
ReDim arrDOC_Col(arrDoc_Nb) As Long

'K = mCols
'Do While Not rsSab.EOF
'    K = K - 1
'    arrDOC_Col(Val(rsSab("BIATABK2"))) = K
'    rsSab.MoveNext
'Loop

 X = ""
 For K = 1 To mCols
     X = X & Format(K, "### ") & "   |"
Next K
fgParam_Recap.FormatString = "Intitulé                                                                                                                          |" & X
 
 
X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'CREDOC' and BIATABK1 = 'Courrier_Doc' order by BIATABTXT"
      
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
      & " where BIATABID = 'CREDOC' and BIATABK1 = 'Courrier_DDS' order by BIATABK2"
      
 Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF

    fgParam_Recap.Rows = fgParam_Recap.Rows + 1
    fgParam_Recap.Row = fgParam_Recap.Rows - 1
    fgParam_Recap.Col = 0: fgParam_Recap.Text = " " & Trim(rsSab("BIATABTXT"))
    arrDDS_Row(Val(rsSab("BIATABK2"))) = fgParam_Recap.Row
    rsSab.MoveNext
Loop


X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'CREDOC' and BIATABK1 = 'Courrier_Des' order by BIATABK2"
      
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
'fgParam_Recap.FormatString = fgParam_Recap_FormatString
fgParam_Recap.Row = 0

currentAction = "fgParam_Recap_Display"
X = "select count(*)  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'CREDOC' and BIATABK1 = 'Courrier_Doc' "
      
 Set rsSab = cnsab.Execute(X)
 mCols = rsSab(0) + 2
 
 X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'CREDOC' and BIATABK1 = 'Courrier_Doc' " _
      & " order by BIATABK2 desc"
      
 Set rsSab = cnsab.Execute(X)
 
arrDoc_Nb = Val(rsSab("BIATABK2"))
ReDim arrDOC_Col(arrDoc_Nb) As Long
'K = mCols
'Do While Not rsSab.EOF
'    K = K - 1
'    arrDOC_Col(Val(rsSab("BIATABK2"))) = K
'    rsSab.MoveNext
'Loop

 X = ""
 For K = 2 To mCols
    X = X & Format(K, "### ") & "   |"
Next K
 fgParam_Recap.FormatString = "Code                           |" _
     & "Intitulé                                                                                                                          |" & X
 
 
X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'CREDOC' and BIATABK1 = 'Courrier_Doc' order by BIATABTXT"
      
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
      & " where BIATABID = 'CREDOC' and BIATABK1 = '#SAB'"
      
 Set rsSab = cnsab.Execute(X)
arrSAB_Nb = rsSab(0) + 1
ReDim arrSAB_Row(arrSAB_Nb) As Long, arrSAB_Id(arrSAB_Nb) As String

X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'CREDOC' and BIATABK1 = '#SAB' order by BIATABK2"
      
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
      & " where BIATABID = 'CREDOC_#SAB'  order by BIATABK1 , BIATABK2"
      
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
    'SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub fgZCDOMOD0_Display()
Dim wColor As Long, X As String, curX As Currency

Dim kRow As Long, kCol As Long

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgZCDOMOD0.Visible = False: lstMT707.Visible = False: lstMT799.Visible = False: lstMT734.Visible = False
fgZCDOMOD0_Reset

fgZCDOMOD0.Rows = 1
fgZCDOMOD0.FormatString = fgZCDOMOD0_FormatString
fgZCDOMOD0.Row = 0

currentAction = "fgZCDOMOD0_Display"
X = "select *  from " & paramIBM_Library_SAB & ".ZCDOMOD0 " _
      & " where CDOMODETB = " & xZCDODOS0.CDODOSETB & " and CDOMODAGE = " & xZCDODOS0.CDODOSAGE _
      & " and CDOMODSER = '" & xZCDODOS0.CDODOSSER & "' and CDOMODSSE = '" & xZCDODOS0.CDODOSSSE & "'" _
      & " and CDOMODCOP = '" & xZCDODOS0.CDODOSCOP & "' and CDOMODDOS = " & xZCDODOS0.CDODOSDOS _
      & " order by CDOMODNMO"
      
 Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF

         fgZCDOMOD0.Rows = fgZCDOMOD0.Rows + 1
         fgZCDOMOD0.Row = fgZCDOMOD0.Rows - 1
        
         fgZCDOMOD0.Col = 0: fgZCDOMOD0.Text = rsSab("CDOMODNMO")
         fgZCDOMOD0.Col = 1: fgZCDOMOD0.Text = " " & dateImp10_S(rsSab("CDOMODDMO") + 19000000)
         fgZCDOMOD0.Col = 2: fgZCDOMOD0.Text = " " & rsSab("CDOMODCON")
         fgZCDOMOD0.Col = 3: fgZCDOMOD0.Text = " " & dateImp10_S(rsSab("CDOMODVAL") + 19000000)
         fgZCDOMOD0.Col = 4: fgZCDOMOD0.Text = " " & dateImp10_S(rsSab("CDOMODDLE") + 19000000)
         fgZCDOMOD0.Col = 5: fgZCDOMOD0.Text = Format(rsSab("CDOMODMON"), "### ### ### ##0.00")
         fgZCDOMOD0.Col = 6: fgZCDOMOD0.Text = Format(rsSab("CDOMODMOC"), "### ### ### ##0.00")
         curX = rsSab("CDOMODMOT") - rsSab("CDOMODMOC") - rsSab("CDOMODMOD")
         fgZCDOMOD0.Col = 7: fgZCDOMOD0.Text = Format(curX, "### ### ### ##0.00")
         fgZCDOMOD0.Col = 8: fgZCDOMOD0.Text = Format(rsSab("CDOMODMOD"), "### ### ### ##0.00")
    rsSab.MoveNext
Loop

'fgZCDOMOD0.Visible = True
If fgZCDOMOD0.Rows > 1 Then
         fgZCDOMOD0.Rows = fgZCDOMOD0.Rows + 1
         fgZCDOMOD0.Row = fgZCDOMOD0.Rows - 1
    fgZCDOMOD0.Col = 0: fgZCDOMOD0.Text = ""
    fgZCDOMOD0.Col = 2: fgZCDOMOD0.Text = " " & xZCDODOS0.CDODOSCON
    fgZCDOMOD0.CellForeColor = vbBlue
    fgZCDOMOD0.Col = 3: fgZCDOMOD0.Text = " " & dateImp10_S(xZCDODOS0.CDODOSVAL + 19000000)
    fgZCDOMOD0.CellForeColor = vbBlue
    fgZCDOMOD0.Col = 4: fgZCDOMOD0.Text = " " & dateImp10_S(xZCDODOS0.CDODOSDLE + 19000000)
    fgZCDOMOD0.CellForeColor = vbBlue
    fgZCDOMOD0.Col = 5: fgZCDOMOD0.Text = Format(xZCDODOS0.CDODOSMON, "### ### ### ##0.00")
    fgZCDOMOD0.CellForeColor = vbBlue
    fgZCDOMOD0.Col = 6: fgZCDOMOD0.Text = Format(xZCDODOS0.CDODOSMOC, "### ### ### ##0.00")
    fgZCDOMOD0.CellForeColor = vbBlue
    curX = xZCDODOS0.CDODOSMOT - xZCDODOS0.CDODOSMOC - xZCDODOS0.CDODOSMOD
    fgZCDOMOD0.Col = 7: fgZCDOMOD0.Text = Format(curX, "### ### ### ##0.00")
    fgZCDOMOD0.Col = 8: fgZCDOMOD0.Text = Format(xZCDODOS0.CDODOSMOD, "### ### ### ##0.00")
    optCourrier_MOD.Visible = True
    For kRow = 2 To fgZCDOMOD0.Rows - 1
        
        For kCol = 2 To 8
            fgZCDOMOD0.Col = kCol
            fgZCDOMOD0.Row = kRow: X = Trim(fgZCDOMOD0.Text)
            fgZCDOMOD0.Row = kRow - 1
            If X <> Trim(fgZCDOMOD0.Text) Then
                fgZCDOMOD0.CellBackColor = mColor_W1
                fgZCDOMOD0.Row = kRow: fgZCDOMOD0.CellBackColor = mColor_Y2
            End If
        Next kCol
        
    Next kRow
    
        fgZCDOMOD0.Row = fgZCDOMOD0.Rows - 1
        Call fgZCDOMOD0_MouseDown(1, 0, 1000, 400)
        'fgZCDOMOD0.TopRow = fgZCDOMOD0.Row
        optCourrier_MOD.value = True

Else
    If xZCDODOS0.CDODOSEVE = "07" Then
        optCourrier_MOD.Visible = True
    Else
        optCourrier_MOD.Visible = False
    End If
    
End If

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgZCDOMOD0.Row): DoEvents


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub lstCourrier_Load()
Dim X As String, K As Integer, K1 As Integer, blnOk As Boolean, blnDisplay As Boolean
Dim wCourrier_Des_Len As Integer, blnDisplay_All As Boolean
Static blnOrderBy As Boolean

On Error Resume Next
'Dim xSelect As String, xSelect2 As String

'Dim arrCourrier_Doc() As typeYBIATAB0, arrCourrier_Doc_Nb As Integer
'Dim arrCourrier_Des() As typeYBIATAB0
If blnOrderBy <> optCourrier_All Then
    blnOrderBy = optCourrier_All
    arrCourrier_Doc_Nb = 0
End If


If arrCourrier_Doc_Nb = 0 Then

    X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
       & " where BIATABID = 'CREDOC' and BIATABK1 = 'Courrier_DDS' and BIATABK2 < '123' order by BIATABK2 Desc"
    Set rsSab = cnsab.Execute(X)
    If Not rsSab.EOF Then mCourrier_Des_Len = Val(rsSab("BIATABK2"))

    X = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
       & " where BIATABID = 'CREDOC' and BIATABK1 = 'Courrier_Des'"

    Set rsSab = cnsab.Execute(X)
    If Not rsSab.EOF Then
        ReDim arrCourrier_Doc(rsSab(0) + 1), arrCourrier_Des(rsSab(0) + 1), arrCourrier_Id(rsSab(0) + 1) _
            , arrCourrier_Originaux_Param_Nb(rsSab(0) + 1), arrCourrier_Originaux_Dossier_Nb(rsSab(0) + 1)
    End If
    X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 A , " & paramIBM_Library_SABSPE & ".YBIATAB0 B" _
       & " where A.BIATABID = 'CREDOC' and A.BIATABK1 = 'Courrier_Des'" _
       & " and B.BIATABID = 'CREDOC' and B.BIATABK1 = 'Courrier_Doc'" _
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

If xZCDODOS0.CDODOSCOP = "CDE" Then
    Mid$(mCourrier_Des, 1, 1) = "O"
Else
    Mid$(mCourrier_Des, 2, 1) = "O"
End If
If optLangue_FR.value = True Then Mid$(mCourrier_Des, 3, 1) = "O"
If optLangue_GB.value = True Then Mid$(mCourrier_Des, 4, 1) = "O"
If optCourrier_OUV.value = True Then Mid$(mCourrier_Des, 6, 1) = "O"
If optCourrier_MOD.value = True Then Mid$(mCourrier_Des, 7, 1) = "O"
If optCourrier_UTI.value = True Then Mid$(mCourrier_Des, 8, 1) = "O"
If optCourrier_All.value = True Then wCourrier_Des_Len = 4: blnDisplay_All = True
Select Case xZCDODOS0.CDODOSCON
    Case "C": Mid$(mCourrier_Des, 9, 1) = "O"
    Case "N": Mid$(mCourrier_Des, 10, 1) = "O"
    Case "P": Mid$(mCourrier_Des, 11, 1) = "O"
End Select

If xZCDODOS0.CDODOSPMO <> 0 Or xZCDODOS0.CDODOSPPO <> 0 Then
    Mid$(mCourrier_Des, 12, 1) = "O"
End If

If xZCDODOS0.CDODOSEVE = "80" Or xZCDODOS0.CDODOSEVE = "90" Then
    Mid$(mCourrier_Des, 13, 1) = "O"
Else
    Mid$(mCourrier_Des, 13, 1) = "N"
End If

If xZCDODOS0.CDODOSMOV <> 0 Then
    Mid$(mCourrier_Des, 14, 1) = "O"
Else
    Mid$(mCourrier_Des, 14, 1) = "N"
End If

If xZCDODOS0.CDODOSMDI <> 0 Then
    Mid$(mCourrier_Des, 15, 1) = "O"
Else
    Mid$(mCourrier_Des, 15, 1) = "N"
End If

If mECNF.WCDOCO2TX1 > 0 Or mENOTIF.WCDOCO2TX1 > 0 Then
    Mid$(mCourrier_Des, 16, 1) = "NO"
Else
    Mid$(mCourrier_Des, 16, 1) = "ON"
End If

If xZCDODOS0.CDODOSNOT = "0011074" Then
    Mid$(mCourrier_Des, 20, 1) = "O"
Else
    Mid$(mCourrier_Des, 20, 1) = "N"
End If

If xZCDODOS0.CDODOSNOT = "0050500" Then
    Mid$(mCourrier_Des, 19, 1) = "O"
Else
    Mid$(mCourrier_Des, 19, 1) = "N"
End If

If xZCDODOS0.CDODOSPAR <> "0010000" Then
    Mid$(mCourrier_Des, 21, 1) = "O"
Else
    Mid$(mCourrier_Des, 21, 1) = "N"
End If

If xZCDODOS0.CDODOSBEN = "0050564" Then
    Mid$(mCourrier_Des, 22, 1) = "O"
Else
    Mid$(mCourrier_Des, 22, 1) = "N"
End If

If xZCDODOS0.CDODOSBEN = "0050517" Then
    Mid$(mCourrier_Des, 23, 1) = "O"
Else
    Mid$(mCourrier_Des, 23, 1) = "N"
End If

If xZCDODOS0.CDODOSBEN = "0050869" Then
    Mid$(mCourrier_Des, 24, 1) = "O"
Else
    Mid$(mCourrier_Des, 24, 1) = "N"
End If

'_________________________________________________________________________________
If blnZCDOUTI0_Select Then

    Mid$(mCourrier_Des, 30, 1) = xZCDOUTI0.CDOUTIDCO
    
    If xZCDOUTI0.CDOUTITMO = "C" Then
        Mid$(mCourrier_Des, 31, 1) = "O"
    Else
        Mid$(mCourrier_Des, 31, 1) = "N"
    End If
    
    If xZCDOUTI0.CDOUTITMO = "N" Then
        Mid$(mCourrier_Des, 32, 1) = "O"
    Else
        Mid$(mCourrier_Des, 32, 1) = "N"
    End If
   
    If xZCDOUTI0.CDOUTITMO = "D" Then
        Mid$(mCourrier_Des, 33, 1) = "O"
    Else
        Mid$(mCourrier_Des, 33, 1) = "N"
    End If
    
    If xZCDOUTI0.CDOUTIMVU > 0 Then
        Mid$(mCourrier_Des, 34, 1) = "O"
    Else
        Mid$(mCourrier_Des, 34, 1) = "N"
    End If
    
    If xZCDOUTI0.CDOUTIMDI > 0 Then
        Mid$(mCourrier_Des, 35, 1) = "O"
    Else
        Mid$(mCourrier_Des, 35, 1) = "N"
    End If
    
    If xZCDOUTI0.CDOUTIEVE = "04" Then
        Mid$(mCourrier_Des, 36, 1) = "O"
    Else
        Mid$(mCourrier_Des, 36, 1) = "N"
    End If
    
    If xZCDODOS0.CDODOSBEC = "O" Then
        Mid$(mCourrier_Des, 40, 1) = "O"
    Else
        Mid$(mCourrier_Des, 40, 1) = "N"
    End If
End If
'_________________________________________________________________________________
If blnZCDOMOD0_Select Then
    
    If mOLD_MTD < mNEW_MTD Then
        Mid$(mCourrier_Des, 41, 1) = "O"
    Else
        Mid$(mCourrier_Des, 41, 1) = "N"
    End If
    If mOLD_MTD > mNEW_MTD Then
        Mid$(mCourrier_Des, 42, 1) = "O"
    Else
        Mid$(mCourrier_Des, 42, 1) = "N"
    End If
    
    If mOLD_MT_C < mNEW_MT_C Then
        Mid$(mCourrier_Des, 43, 1) = "O"
    Else
        Mid$(mCourrier_Des, 43, 1) = "N"
    End If
    If mOLD_MT_N < mNEW_MT_N Then
        Mid$(mCourrier_Des, 44, 1) = "O"
    Else
        Mid$(mCourrier_Des, 44, 1) = "N"
    End If
    
    If mOLD_CND = "C" And mNEW_CND = "N" Then
        Mid$(mCourrier_Des, 45, 1) = "O"
    Else
        Mid$(mCourrier_Des, 45, 1) = "N"
    End If
    If mOLD_CND = "N" And mNEW_CND = "C" Then
        Mid$(mCourrier_Des, 46, 1) = "O"
    Else
        Mid$(mCourrier_Des, 46, 1) = "N"
    End If
     
    If mOLD_VALIDIT < mNEW_VALIDIT Then
        Mid$(mCourrier_Des, 47, 1) = "O"
    Else
        Mid$(mCourrier_Des, 47, 1) = "N"
    End If
     
    If mOLD_EMB < mNEW_EMB Then
        Mid$(mCourrier_Des, 48, 1) = "O"
    Else
        Mid$(mCourrier_Des, 48, 1) = "N"
    End If
     
    If mOLD_VALIDIT > mNEW_VALIDIT Then
        Mid$(mCourrier_Des, 49, 1) = "O"
    Else
        Mid$(mCourrier_Des, 49, 1) = "N"
    End If
     
    If mOLD_EMB > mNEW_EMB Then
        Mid$(mCourrier_Des, 50, 1) = "O"
    Else
        Mid$(mCourrier_Des, 50, 1) = "N"
    End If
   
End If
'_________________________________________________________________________________

lstCourrier.Clear

For K = 1 To arrCourrier_Doc_Nb
    blnOk = True: blnDisplay = True
    'If arrCourrier_Id(K) = "000000000140" Then
    '    Debug.Print arrCourrier_Doc(K)
    'End If
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
                'Case " ":
                '    Select Case Mid$(mCourrier_Des, K1, 1)
                '        Case " ":
                '        Case Else:
                '            blnOk = False
                '    End Select
            End Select
           
        Next K1
    If Mid$(arrCourrier_Des(K), 123, 1) = "O" Then blnOk = False
    If blnDisplay_All Then blnOk = False
    If Mid$(arrCourrier_Des(K), 124, 1) = "O" Then
        X = "+ " & arrCourrier_Doc(K)
    Else
         X = "  " & arrCourrier_Doc(K)
   End If
    
    End If
    If blnOk Then
        lstCourrier.AddItem X
        lstCourrier.Selected(lstCourrier.ListCount - 1) = True
    Else
        If blnDisplay Then lstCourrier.AddItem X
    End If
Next K

'lstCourrier.Visible = True
lstCourrier.Visible = arrHab(2)
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


Public Sub fgZCDOMOD0_Reset()
fgZCDOMOD0.Clear
fgZCDOMOD0_Sort1 = 0: fgZCDOMOD0_Sort2 = 0
fgZCDOMOD0_Sort1_Old = -1
fgZCDOMOD0_RowDisplay = 0: fgZCDOMOD0_RowClick = 0
fgZCDOMOD0_arrIndex = fgZCDOMOD0.Cols - 1
blnfgZCDOMOD0_DisplayLine = False
fgZCDOMOD0_SortAD = 6
fgZCDOMOD0.LeftCol = fgZCDOMOD0.FixedCols

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


Public Sub fgUTI_DOC_Sort()
If fgUTI_DOC.Rows > 1 Then
    fgUTI_DOC.Row = 1
    fgUTI_DOC.RowSel = fgUTI_DOC.Rows - 1
    
    fgUTI_DOC.Col = 3
    fgUTI_DOC.ColSel = 3
    fgUTI_DOC.Sort = 5
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


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim xSql As String

If y <= fgSelect.RowHeightMin Then
Else
    If fgSelect.Rows > 1 And y < fgSelect.Rows * fgSelect.CellHeight Then
        blnUTI_Com_Loaded = False: blnUTI_Com_Ok = False: fgUTI_Com.Visible = False
        fgUTI_DOC.Visible = False
        fgSelect.Col = 0: xZCDOUTI0.CDOUTIUTI = fgSelect.Text
        fgSelect.Col = 5: xZCDOUTI0.CDOUTINUR = fgSelect.Text
        
        xSql = "select *  from " & paramIBM_Library_SAB & ".ZCDOUTI0" _
              & " where CDOUTIETB = " & xZCDODOS0.CDODOSETB & " and CDOUTIAGE = " & xZCDODOS0.CDODOSAGE _
              & " and CDOUTISER = '" & xZCDODOS0.CDODOSSER & "' and CDOUTISSE = '" & xZCDODOS0.CDODOSSSE & "'" _
              & " and CDOUTICOP = '" & xZCDODOS0.CDODOSCOP & "' and CDOUTIDOS = " & xZCDODOS0.CDODOSDOS _
              & " and CDOUTINUR = '" & xZCDOUTI0.CDOUTINUR & "' and CDOUTIUTI = " & xZCDOUTI0.CDOUTIUTI
        Set rsSab = cnsab.Execute(xSql)
        
        If Not rsSab.EOF Then
    
            Call rsZCDOUTI0_GetBuffer(rsSab, xZCDOUTI0)
            
            lstCourrier_Load
            
            If xZCDOUTI0.CDOUTIBEC = "O" Then
                mUTI_BEC = "Nos frais et ceux de la banque émettrice seront déduits lors du règlement."
            Else
                 mUTI_BEC = ""
            End If
            
            Call txtRTF_ZCDOIRR0
            blnZCDOUTI0_Select = True
            optCourrier_UTI.Visible = True
            optCourrier_UTI = True
            If Trim(xZCDOUTI0.CDOUTIREM) = "" Then
                Call rsZADRESS0_Init(mREM_ZADRESS0)
                mREM_Concat = ""
            Else
                xSql = Space(64)
                'DR 01/06/2017
                'Call rsZCDOTIE_Adresse(xZCDOUTI0.CDOUTICTR, xZCDOUTI0.CDOUTIREM, xSql, xYCDOTIE0, mREM_ZADRESS0, mREM_Concat, "  ")
                Call rsZCDOTIE_Adresse(xZCDOUTI0.CDOUTICTR, xZCDOUTI0.CDOUTIREM, xSql, xYCDOTIE0, mREM_ZADRESS0, mREM_Concat, "CO")
                If Trim(mREM_ZADRESS0.ADRESSPAY) = "FRANCE" Then mREM_ZADRESS0.ADRESSPAY = ""
            End If
        End If
        
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)

    End If
End If

End Sub

Private Sub fgUTI_Com_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next

If y <= fgUTI_Com.RowHeightMin Then
Else
    If fgUTI_Com.Rows > 1 And fgUTI_Com.Row > mUTI_Com_RowMin Then 'And X > 3930 Then
        If fgUTI_Com.CellBackColor = 0 Or fgUTI_Com.CellBackColor = vbWhite Or fgUTI_Com.CellBackColor = mColor_G1 Then
            Select Case X
                Case Is < 4530: fgUTI_Com.Col = 0: mUTI_Com_Col = 0
                Case Is < 6330: fgUTI_Com.Col = 1: mUTI_Com_Col = 1
                Case Is < 7680: fgUTI_Com.Col = 2: mUTI_Com_Col = 2
                Case Else: fgUTI_Com.Col = 3: mUTI_Com_Col = 3
            End Select
            If mUTI_Com_Col < 3 Then
                fgUTI_Com.ScrollBars = flexScrollBarNone
                txtUTI_Com_M.Top = fgUTI_Com.CellTop + fgUTI_Com.CellHeight * 2 + 80
                txtUTI_Com_M.Height = fgUTI_Com.RowHeight(fgUTI_Com.Row)
                txtUTI_Com_M.Left = fgUTI_Com.CellLeft + fgUTI_Com.Left
                txtUTI_Com_M.Width = fgUTI_Com.CellWidth
                txtUTI_Com_M.Text = Trim(fgUTI_Com.Text)
                
                fgUTI_Com.Col = 0: fgUTI_Com.CellBackColor = fgUTI_Com.BackColorSel '&H80C0FF
                txtUTI_Com_M.Visible = True
                txtUTI_Com_M.SetFocus
                txtUTI_Com_M.ZOrder 0
            End If
        End If
    End If
End If

End Sub

Private Sub fgUTI_DOC_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next

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
'fgUTI_DOC.RowPos
End Sub


Private Sub fgZCDOMOD0_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim xSql As String

If y <= fgZCDOMOD0.RowHeightMin Then
Else
    If fgZCDOMOD0.Rows > 1 And fgZCDOMOD0.Row > 1 Then
        'fgZCDOMOD0.Col = 0: xZCDOMOD0.CDOMODNMO = Val(fgZCDOMOD0.Text)
        Call fgZCDOMOD0_Color(fgZCDOMOD0_RowClick, MouseMoveUsr.BackColor, fgZCDOMOD0_ColorClick)
        
        fgZCDOMOD0.Col = 2: mNEW_CND = Trim(fgZCDOMOD0.Text)
        fgZCDOMOD0.Col = 3: Call dateJMA_AMJ(Trim(fgZCDOMOD0.Text), mNEW_VALIDIT)
        fgZCDOMOD0.Col = 4: Call dateJMA_AMJ(Trim(fgZCDOMOD0.Text), mNEW_EMB)
        fgZCDOMOD0.Col = 5: mNEW_MTD = Trim(fgZCDOMOD0.Text)
        fgZCDOMOD0.Col = 6: mNEW_MT_C = Trim(fgZCDOMOD0.Text)
        fgZCDOMOD0.Col = 7: mNEW_MT_N = Trim(fgZCDOMOD0.Text)
        fgZCDOMOD0.Col = 8: mNEW_MT_D = Trim(fgZCDOMOD0.Text)
    
        fgZCDOMOD0.Row = fgZCDOMOD0.Row - 1
        fgZCDOMOD0.Col = 2: mOLD_CND = Trim(fgZCDOMOD0.Text)
        fgZCDOMOD0.Col = 3: Call dateJMA_AMJ(Trim(fgZCDOMOD0.Text), mOLD_VALIDIT)
        fgZCDOMOD0.Col = 4: Call dateJMA_AMJ(Trim(fgZCDOMOD0.Text), mOLD_EMB)
        fgZCDOMOD0.Col = 5: mOLD_MTD = Trim(fgZCDOMOD0.Text)
        fgZCDOMOD0.Col = 6: mOLD_MT_C = Trim(fgZCDOMOD0.Text)
        fgZCDOMOD0.Col = 7: mOLD_MT_N = Trim(fgZCDOMOD0.Text)
        fgZCDOMOD0.Col = 8: mOLD_MT_D = Trim(fgZCDOMOD0.Text)
       'xSql = "select *  from " & paramIBM_Library_SAB & ".ZCDOMOD0" _
        '      & " where CDOMODETB = " & xZCDODOS0.CDODOSETB & " and CDOMODAGE = " & xZCDODOS0.CDODOSAGE _
        '      & " and CDOMODSER = '" & xZCDODOS0.CDODOSSER & "' and CDOMODSSE = '" & xZCDODOS0.CDODOSSSE & "'" _
        '      & " and CDOMODCOP = '" & xZCDODOS0.CDODOSCOP & "' and CDOMODDOS = " & xZCDODOS0.CDODOSDOS _
        '      & " and CDOMODNMO = " & xZCDOMOD0.CDOMODNMO
        'Set rsSab = cnsab.Execute(xSql)
        
        'If Not rsSab.EOF Then
    
         '   Call rsZCDOMOD0_GetBuffer(rsSab, xZCDOMOD0)
            
            
            blnZCDOMOD0_Select = True
            optCourrier_MOD.Visible = True
            optCourrier_MOD = True
            lstCourrier_Load

            
        'End If
        

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
If fgZCDOMOD0.Visible Then
    fgZCDOMOD0.Visible = False: lstMT707.Visible = False: lstMT799.Visible = False: lstMT734.Visible = False
    optCourrier_OUV = True: Exit Sub
End If
If txtUTI_DOC_M.Visible Then txtUTI_DOC_M.Visible = False: Exit Sub
If fgUTI_DOC.Visible Then fgUTI_DOC.Visible = False: txtUTI_DOC_M.Visible = False: Exit Sub
If txtUTI_Com_M.Visible Then txtUTI_Com_M.Visible = False: Exit Sub
If fgUTI_Com.Visible Then fgUTI_Com.Visible = False: txtUTI_Com_M.Visible = False: Exit Sub
If fraParam_Courrier.Visible Then fraParam_Courrier.Visible = False: Exit Sub
If txtInfo_M.Visible Then txtInfo_M.Visible = False: Exit Sub
If fraInfo_M.Visible Then Call cmdInfo_M_Quit_Click: Exit Sub  'fraInfo_M.Visible = False: Exit Sub
'If txtRTF.Visible Then txtRTF.Visible = False: Exit Sub

If txtFg.Visible Then txtFg.Visible = False: Exit Sub

Unload Me

End Sub

Private Sub Form_Load()

frmSAB_Dossier_CDO_Show

Set XForm = Me
Me.Left = 19000 - Me.Width
KeyPreview = True

blnControl = False
'mWindowState = Me.WindowState
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate


fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False

fgZCDOMOD0.Visible = False: lstMT707.Visible = False: lstMT799.Visible = False: lstMT734.Visible = False
fgZCDOMOD0_FormatString = fgZCDOMOD0.FormatString
fgZCDOMOD0.Top = fgSelect.Top
fgZCDOMOD0.Left = fgSelect.Left

fgParam_Courrier_FormatString = fgParam_Courrier.FormatString
fgParam_Courrier.Enabled = True
fgParam_Courrier.Visible = False

fgParam_Courrier_FormatString = fgParam_Recap.FormatString
fgParam_Recap.Enabled = True
fgParam_Recap.Visible = False

SSTab1.Tab = 0
Set fraInfo_M.Container = fraDossier
fraInfo_M.Top = txtRTF.Top
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

fgUTI_Com.Visible = False: txtUTI_Com_M.Visible = False
fgUTI_Com.Top = fgInfo_M.Top
fgUTI_Com.Left = fgInfo_M.Left + fgInfo_M.Width - fgUTI_Com.Width


lstMT707.Visible = False

End Sub


Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub


Public Sub Form_Init(lFct As String, lCDODOSCOP As String, lCDODOSDOS As Long)
'___________________________________________________________
Dim X As String, K As Integer
On Error Resume Next
Call BIA_VB_HAB("SAB_DOS_CDO", arrHab(), cboSelect_SQL)
If Not arrHab(1) Then Unload Me: Exit Sub

sstabParam.Tab = 0
SSTab1.Tab = 0

chkWord_Update.Visible = arrHab(3)
chkWord_Validation.Visible = arrHab(2)
chkWord_Validation = "1"
chkPDF_Display.Visible = False 'arrHab(2)
lstCourrier.Visible = False
sstabParam.Visible = arrHab(18)
fraInfo_M.Visible = False
'_______________________________________________________________________________________________________
X = "select count(*)  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'CREDOC' and BIATABK1 = '#SAB'"
      
Set rsSabX = cnsab.Execute(X)
If Not rsSabX.EOF Then
    ReDim arrFields_SAB_Name(rsSabX(0) + 10), arrFields_SAB_Value(rsSabX(0) + 10), blnFields_SAB_Name(rsSabX(0) + 10)
    arrFields_SAB_Nb = 0
    X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
          & " where BIATABID = 'CREDOC' and BIATABK1 = '#SAB' order by BIATABK2"
          
    Set rsSabX = cnsab.Execute(X)
    Do Until rsSabX.EOF
        arrFields_SAB_Nb = arrFields_SAB_Nb + 1
        arrFields_SAB_Name(arrFields_SAB_Nb) = Trim(rsSabX("BIATABK2"))
        rsSabX.MoveNext
    Loop
End If
'_______________________________________________________________________________________________________
X = "select count(*)  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
      & " where BIATABID = 'CREDOC' and BIATABK1 = '?BIA'"
      
Set rsSabX = cnsab.Execute(X)
If Not rsSabX.EOF Then
    ReDim arrFields_BIA_Name(rsSabX(0) + 10), arrFields_BIA_Value(rsSabX(0) + 10), blnFields_BIA_Name(rsSabX(0) + 10), arrFields_BIA_Lib(rsSabX(0) + 10)
    arrFields_BIA_Nb = 0
    X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
          & " where BIATABID = 'CREDOC' and BIATABK1 = '?BIA' order by BIATABK2"
          
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
    If paramCREDOC_TEL_NEGO = "" Then paramCREDOC_Init
End If
cmdPrint.Visible = arrHab(2)


If Not frmSAB_Dossier_CDO.Visible Then frmSAB_Dossier_CDO.Visible = True
Me.Show

fgSelect.Visible = False
txtRTF.Visible = False
mFct_Caller = lFct
'rsZCDODOS0_Init xZCDODOS0

X = "select *  from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
      & " where CDODOSETB = " & currentSAB_ETA & " and CDODOSAGE = " & currentSAB_AGE _
      & " and CDODOSSER = '00' and CDODOSSSE = '00'" _
      & " and CDODOSCOP = '" & lCDODOSCOP & "' and CDODOSDOS = " & lCDODOSDOS
      
 Set rsSabX = cnsab.Execute(X)
If rsSabX.EOF Then
    Call MsgBox("Dossier inconnu : " & lCDODOSCOP & " " & lCDODOSDOS, vbCritical, "SAB_DOssier_CDO")
    Exit Sub
Else
    Call lstErr_Clear(lstErr, cmdContext, "Lecture du dossier : " & lCDODOSCOP & " " & lCDODOSDOS): DoEvents
    lstErr.Height = 510
    Call rsZCDODOS0_GetBuffer(rsSabX, xZCDODOS0)
    Call lstErr_AddItem(lstErr, cmdContext, "Affichage du dossier "): DoEvents
    Call fraDossier_Display
    If arrHab(2) Then lstCourrier_Load
End If

Call lstPrinters_Load
'


If arrHab(2) Then
    K = Windows_Processus_Actif("WINWORD")
    If K > 0 Then
           Call MsgBox("Attention, il y a déjà " & K & " instance(s) 'Word' active(s)!" & vbCrLf & vbCrLf & "Veuillez fermer les documents 'Word', si vous devez éditer des courriers", vbExclamation, "SAB_Dossier : courrier")
    End If

End If

X = Trim(frmSAB_Dossier_CDO.Caption)
AppActivate X

End Sub


Public Sub txtRTF_ZCDODOS0()
Dim X As String
txtRTF.LoadFile (paramCDO_Dossier_Path_DROPI & "Modèles\" & "ZCDODOS0.rtf")
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#COP _DOS", mCOP_DOS)
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#CON", mCON)
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#EVE", mEVE)
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#ETA", mETA)

If xZCDODOS0.CDODOSEVE = "80" Or xZCDODOS0.CDODOSEVE = "90" Then
    If xZCDODOS0.CDODOSDAN = 0 Then
        txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#CDODOSDAN", "")
    Else
        txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#CDODOSDAN", "\cf2\highlight1 annulé le :" & " \cf0\highlight0  " & dateImp10_S(xZCDODOS0.CDODOSDAN + 19000000))
    End If
    
    If xZCDODOS0.CDODOSCLO = 0 Then
        txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#CDODOSCLO", "")
    Else
        txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#CDODOSCLO", "\cf2\highlight1 clos le :" & " \cf0\highlight0  " & dateImp10_S(xZCDODOS0.CDODOSCLO + 19000000))
    End If
Else
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#CDODOSDAN", "")
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#CDODOSCLO", "")
End If

txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#EXT", Trim(xZCDODOS0.CDODOSEXT))
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#MON_DEV", mMON_DEV)
If xZCDODOS0.CDODOSMOV <> 0 Then
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#MOV", Format(xZCDODOS0.CDODOSMOV, "### ### ### ##0.00"))
Else
    'txtRTF.TextRTF = Replace(txtRTF.TextRTF, "\tab\cf0\b0 A Vue \tab :\cf1\b  #MOV ", "")
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#MOV", "-")
End If
If xZCDODOS0.CDODOSMDI <> 0 Then
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#MDI", Format(xZCDODOS0.CDODOSMDI, "### ### ### ##0.00"))
Else
    'txtRTF.TextRTF = Replace(txtRTF.TextRTF, "\tab\cf0\b0 P.Dif :\cf1\b  #MDI", "")
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#MDI", "-")
End If
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#OUV_VAL", mOUV_VAL)
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#X_OUI", mX_OUI)
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#X_NON", mX_NON)

'_________________________________________________________________________________________________________________

If xZCDODOS0.CDODOSCOP = "CDE" Then
    w_ZADRESSE0 = mBQE_ZADRESS0
    Call txtRTF_ZCDODOS0_ADR("#BQE_RS", "#BQE_*")
Else
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "Banque \'e9mettrice", "Banque notificatrice")
    w_ZADRESSE0 = mNOT_ZADRESS0
    Call txtRTF_ZCDODOS0_ADR("#BQE_RS", "#BQE_*")
End If
'_________________________________________________________________________________________________________________

If blnBED_ZADRESSE0 Then
    w_ZADRESSE0 = mBED_ZADRESS0
    Call txtRTF_ZCDODOS0_ADR("#BED_RS", "#BED_*")
Else
    X = "\par \pard\nowidctlpar\li1418\sl276\slmult1\b0 #BED_*"
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, X, "")
    '$JPL 2013-06-18 X = "\par \cf0\highlight6\b\i Courrier Si\'e8ge       :\highlight0  \b0\i0  \cf3\b #BED_RS"
    X = "\b\i Courrier Si\'e8ge       :\highlight0  \b0\i0  \cf3\b #BED_RS"
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, X, "")
End If
'_________________________________________________________________________________________________________________
w_ZADRESSE0 = mDON_ZADRESS0
Call txtRTF_ZCDODOS0_ADR("#DON_RS", "#DON_*")
'_________________________________________________________________________________________________________________
mBEN_TVANIFCLIT = ""
If xZCDODOS0.CDODOSBER = "T" Then
    X = "select TVANIFCLIT  from " & paramIBM_Library_SABSPE & ".YTVANIF0" _
          & " where TVANIFCLIC = 'D' and TVANIFCLI ='" & xZCDODOS0.CDODOSBEN & "'"
    Set rsSabX = cnsab.Execute(X)
        
    If Not rsSabX.EOF Then mBEN_TVANIFCLIT = Trim(rsSabX("TVANIFCLIT"))
Else
    X = "select CLIFISNIF  from " & paramIBM_Library_SAB & ".ZCLIFIS0" _
          & " where CLIFISETA = 1 and CLIFISCLI ='" & xZCDODOS0.CDODOSBEN & "' and CLIFISTYP = 1"
    Set rsSabX = cnsab.Execute(X)
        
    If Not rsSabX.EOF Then mBEN_TVANIFCLIT = Trim(rsSabX("CLIFISNIF"))
End If


w_ZADRESSE0 = mBEN_ZADRESS0
Call txtRTF_ZCDODOS0_ADR("#BEN_RS", "#BEN_*")

'_________________________________________________________________________________________________________________

txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#DEV", xZCDODOS0.CDODOSDEV)
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#MTD_T", mMTD_T)
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#MTD_N", mMTD_N)
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#MTD_D", mMTD_D)
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#MTD_C", mMTD_C)
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#%_N", mRatio_N)
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#%_C", mRatio_C)

If xZCDODOS0.CDODOSPMO <> 0 Then
    X = Format(xZCDODOS0.CDODOSPMO, "### ### ### ##0.00")
Else
    X = ""
End If
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#MTD_PMO", X)

If xZCDODOS0.CDODOSPPO <> 0 Then
    X = Format(xZCDODOS0.CDODOSPPO, "##0") & " %"
Else
    X = ""
End If
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#%_PPO", X)

'_________________________________________________________________________________________________________________
X = ""
If xZCDODOS0.CDODOSDLE <> 0 Then
    X = "\highlight2date limite d'embarquement :" & dateImp10_S(xZCDODOS0.CDODOSDLE + 19000000) & "\highlight0"
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#DLE\cf3\tab\tab\tab\tab", X)
Else
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#DLE", "")
End If

txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#DOC_P", Trim(xZCDODOS0.CDODOSPDO) & Trim(xZCDODOS0.CDODOSPD2))
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#LIEU_LED", Trim(xZCDODOS0.CDODOSLED))
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#LIEU_LDA", Trim(xZCDODOS0.CDODOSLDA))

'_________________________________________________________________________________________________________________
If xZCDODOS0.CDODOSBEC = "O" Then
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#COM_B/O", "Frais Charge Bénéficiaire")
Else
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#COM_B/O", "Frais Charge Ordonnateur")
End If

txtRTF.Visible = True
End Sub

Public Sub fraDossier_Display()
Dim X As String, curX As Currency, iRatio As Integer, K As Integer, K2 As Integer

optCourrier_MOD.Visible = False
optCourrier_UTI.Visible = False
optCourrier_OUV = True
blnUTI_Com_Loaded = False: fgUTI_Com.Visible = False

'If xZCDODOS0.CDODOSEVE = "01" Then
'    optCourrier_OUV = True
'Else
'    optCourrier_MOD = True
'End If

mCOP_DOS = xZCDODOS0.CDODOSCOP & " " & Format(xZCDODOS0.CDODOSDOS, "### 000")


mMON_DEV = Format(xZCDODOS0.CDODOSMON, "### ### ### ###.00") & " " & xZCDODOS0.CDODOSDEV

Select Case xZCDODOS0.CDODOSCON
    Case "C": mCON = "Confirmé"
    Case "N": mCON = "Notifié"
    Case "P": mCON = "Partiel"
    Case Else: mCON = xZCDODOS0.CDODOSCON
End Select
 Select Case xZCDODOS0.CDODOSEVE
    Case "01": mEVE = "Ouverture"
    Case "02": mEVE = "Modification"
    Case "07": mEVE = "Réouverture"
    Case "80": mEVE = "Annulation"
    Case "90": mEVE = "Clôture"
    Case Else: mEVE = xZCDODOS0.CDODOSEVE
End Select
 Select Case xZCDODOS0.CDODOSETA
    Case "01": mETA = X & "-Saisie"
    Case "02": mETA = X & "-Validée"
    Case "03": mETA = X & "-Comptabilisée"
    Case Else: mETA = X & "-" & xZCDODOS0.CDODOSETA
End Select

mOUV_VAL = dateImp10_S(xZCDODOS0.CDODOSOUV + 19000000) & " - " & dateImp10_S(xZCDODOS0.CDODOSVAL + 19000000)
mX_OUI = ""
mX_NON = ""

'Select Case xZCDODOS0.CDODOSOPE
'    Case "O": mX_OUI = mX_OUI & "Opératif"
'    Case Else: mX_NON = mX_NON & "NON Opératif"
'End Select

'Select Case xZCDODOS0.CDODOSIRR
'    Case "O": mX_NON = mX_NON & " - Irrévocable"
'    Case Else: mX_OUI = mX_OUI & " - Révocable"
'End Select

Select Case xZCDODOS0.CDODOSFRA
    Case "O": mX_OUI = mX_OUI & " - Fractionable"
    Case Else: mX_NON = mX_NON & " - NON Fractionable"
End Select

Select Case xZCDODOS0.CDODOSREN
    Case "O": mX_OUI = mX_OUI & " - Renouvelable"
    Case Else: mX_NON = mX_NON & " - NON Renouvelable"
End Select

Select Case xZCDODOS0.CDODOSCUM
    Case "O": mX_OUI = mX_OUI & " - Cumulatif"
    Case Else: mX_NON = mX_NON & " - NON Cumulatif"
End Select

Select Case xZCDODOS0.CDODOSTRS
    Case "O": mX_OUI = mX_OUI & " - Transférable "
    Case Else: mX_NON = mX_NON & " - NON Transférable "
End Select

Select Case xZCDODOS0.CDODOSEPA
    Case "O": mX_OUI = mX_OUI & " - Expédition partielle "
    Case Else: mX_NON = mX_NON & " - NON Expédition partielle "
End Select

Select Case xZCDODOS0.CDODOSTRA
    Case "O": mX_OUI = mX_OUI & " - Transbordement "
    Case Else: mX_NON = mX_NON & " - NON Transbordement "
End Select


'Lecture BQE agence - Donneur d'ordre - Bénéficiaire
X = Space(64)
'DR 01/06/2017
'Call rsZCDOTIE_Adresse(xZCDODOS0.CDODOSCOT, xZCDODOS0.CDODOSCOR, X, xYCDOTIE0, mBQE_ZADRESS0, mBQE_Concat, "CD")
Call rsZCDOTIE_Adresse(xZCDODOS0.CDODOSCOT, xZCDODOS0.CDODOSCOR, X, xYCDOTIE0, mBQE_ZADRESS0, mBQE_Concat, "CO")
If Trim(mBQE_ZADRESS0.ADRESSPAY) = "FRANCE" Then mBQE_ZADRESS0.ADRESSPAY = ""

'DR 01/06/2017
'Call rsZCDOTIE_Adresse(xZCDODOS0.CDODOSDOR, xZCDODOS0.CDODOSDON, xZCDODOS0.CDODOSDOE, xYCDOTIE0, mDON_ZADRESS0, mDON_Concat, "CD")
Call rsZCDOTIE_Adresse(xZCDODOS0.CDODOSDOR, xZCDODOS0.CDODOSDON, xZCDODOS0.CDODOSDOE, xYCDOTIE0, mDON_ZADRESS0, mDON_Concat, "CO")
If Trim(mDON_ZADRESS0.ADRESSPAY) = "FRANCE" Then mDON_ZADRESS0.ADRESSPAY = ""

'DR 01/06/2017
'Call rsZCDOTIE_Adresse(xZCDODOS0.CDODOSBER, xZCDODOS0.CDODOSBEN, xZCDODOS0.CDODOSBEI, xYCDOTIE0, mBEN_ZADRESS0, mBEN_Concat, "  ")
Call rsZCDOTIE_Adresse(xZCDODOS0.CDODOSBER, xZCDODOS0.CDODOSBEN, xZCDODOS0.CDODOSBEI, xYCDOTIE0, mBEN_ZADRESS0, mBEN_Concat, "CO")
If Trim(mBEN_ZADRESS0.ADRESSPAY) = "FRANCE" Then mBEN_ZADRESS0.ADRESSPAY = ""
mBEN_CDOTIESRN = Trim(xYCDOTIE0.CDOTIESRN)

'Lecture Adresse siège pour courrier Bordereau Envoi Doc si bq émettrice : 00110066 (BNA)
If Trim(xZCDODOS0.CDODOSNOT) = "0011066" Then
    blnBED_ZADRESSE0 = True
    Call rsZCDOTIE_Adresse(xZCDODOS0.CDODOSNOR, xZCDODOS0.CDODOSNOT, X, xYCDOTIE0, mBED_ZADRESS0, mBED_Concat, "  ")
    If Trim(mBED_ZADRESS0.ADRESSPAY) = "FRANCE" Then mBED_ZADRESS0.ADRESSPAY = ""

Else
    blnBED_ZADRESSE0 = False
    mBED_ZADRESS0 = mBQE_ZADRESS0
    mBED_Concat = mBQE_Concat
    
End If

'DR 01/06/2017
'Call rsZCDOTIE_Adresse(xZCDODOS0.CDODOSNOR, xZCDODOS0.CDODOSNOT, X, xYCDOTIE0, mNOT_ZADRESS0, mNOT_Concat, "  ")
Call rsZCDOTIE_Adresse(xZCDODOS0.CDODOSNOR, xZCDODOS0.CDODOSNOT, X, xYCDOTIE0, mNOT_ZADRESS0, mNOT_Concat, "CO")
If Trim(mNOT_ZADRESS0.ADRESSPAY) = "FRANCE" Then mNOT_ZADRESS0.ADRESSPAY = ""

'DR 01/06/2017
'Call rsZCDOTIE_Adresse(xZCDODOS0.CDODOSTBR, xZCDODOS0.CDODOSBRE, X, xYCDOTIE0, mBQE_RBT_ZADRESS0, mBQE_RBT_Concat, "  ")
Call rsZCDOTIE_Adresse(xZCDODOS0.CDODOSTBR, xZCDODOS0.CDODOSBRE, X, xYCDOTIE0, mBQE_RBT_ZADRESS0, mBQE_RBT_Concat, "CO")
If Trim(mBQE_RBT_ZADRESS0.ADRESSPAY) = "FRANCE" Then mBQE_RBT_ZADRESS0.ADRESSPAY = ""




mMTD_T = "": mMTD_C = "": mMTD_D = "": mMTD_N = ""
mRatio_C = "": mRatio_N = ""

mMTD_T = Format(xZCDODOS0.CDODOSMOT, "### ### ### ##0.00")
If xZCDODOS0.CDODOSCON = "N" Then
    If xZCDODOS0.CDODOSMOT <> 0 Then mMTD_N = Format(xZCDODOS0.CDODOSMOT, "### ### ### ##0.00")
Else
    If xZCDODOS0.CDODOSMOD <> 0 Then
        mMTD_D = Format(xZCDODOS0.CDODOSMOD, "### ### ### ##0.00")
    Else
        mMTD_C = Format(xZCDODOS0.CDODOSMOC, "### ### ### ##0.00")
    End If
    If xZCDODOS0.CDODOSCON = "P" Then
        curX = xZCDODOS0.CDODOSMOT - xZCDODOS0.CDODOSMOC - xZCDODOS0.CDODOSMOD
        mMTD_N = Format(curX, "### ### ### ##0.00")
        iRatio = Round(curX / xZCDODOS0.CDODOSMOT * 100, 0)
        mRatio_N = iRatio & " %"
        mRatio_C = (100 - iRatio) & " %"
   End If
End If



'
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
blnZCDOTCO0_CDE = False: blnZCDOTCO0_CDI = False
blnZCDOUTI0_Select = False
blnUTI_DOC_Loaded = False
blnZCDOMOD0_Select = False
mAR_Accord = "accord"
mAR_Courrier = "courrier"
mATTN = "A l'attention de M."
'____________________________________________________________________________________

Call txtRTF_ZCDODOS0
Call txtRTF_ZCDOTC20
Call txtRTF_ZCDOCOM0_Ouverture
Call txtRTF_ZCDODES0

Call txtRTF_YSWISAB0

Call fgZCDOMOD0_Display

Call fgSelect_Display


'If fgselect. Then optCourrier_OUV = True
'_______________________________________________________________________
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
appWord.Quit False
Set appWord = Nothing

End Sub

Private Sub lstCourrier_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wText As String, K1 As Integer, wCourrier_Id As Long

If lstCourrier.Visible And X > 300 Then
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
    Exit Sub
End If
End Sub

Private Sub lstMT734_Click()
Dim X As String, K As Integer
On Error GoTo Error_Handler
lstMT734.Visible = False
X = lstMT734.Text
K = InStr(X, "(")
mSWISABSWID_734 = Val(Mid$(X, K + 1, Len(X) - K - 1))

Exit Sub

Error_Handler:
    Call MsgBox(Error, vbCritical, "SAB_DOS_CDO.lstMT734_Click")

End Sub

Private Sub lstMT799_Click()
Dim X As String, K As Integer
On Error GoTo Error_Handler
lstMT799.Visible = False
X = lstMT799.Text
K = InStr(X, "(")
mSWISABSWID_799 = Val(Mid$(X, K + 1, Len(X) - K - 1))

Exit Sub

Error_Handler:
    Call MsgBox(Error, vbCritical, "SAB_DOS_CDO.lstMT799_Click")


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

Private Sub lstParam_Modèles_CREDOC_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim xSql As String
Me.Enabled = False: Me.MousePointer = vbHourglass
fraParam_Courrier.Visible = False

oldCourrier_Doc.BIATABID = "CREDOC"
oldCourrier_Doc.BIATABK1 = "Courrier_Doc"
oldCourrier_Doc.BIATABK2 = ""
oldCourrier_Doc.BIATABTXT = lstParam_Modèles_CREDOC.Text
blnCourrier_Doc_Exist = False
xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CREDOC'" _
     & " and BIATABK1 = 'Courrier_Doc' and BIATABTXT = '" & Trim(oldCourrier_Doc.BIATABTXT) & "'"
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
    blnCourrier_Doc_Exist = True
    oldCourrier_Doc.BIATABK2 = rsSab("BIATABK2")
End If

oldCourrier_Des.BIATABID = "CREDOC"
oldCourrier_Des.BIATABK1 = "Courrier_Des"
oldCourrier_Des.BIATABK2 = oldCourrier_Doc.BIATABK2
blnCourrier_Des_Exist = False
xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CREDOC'" _
     & " and BIATABK1 = 'Courrier_Des' and BIATABK2 = '" & oldCourrier_Doc.BIATABK2 & "'"
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
    blnCourrier_Des_Exist = True
    oldCourrier_Des.BIATABTXT = rsSab("BIATABTXT")
Else
    oldCourrier_Des.BIATABTXT = ""
End If

Me.PopupMenu mnuParam_Modèles_CREDOC
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

oldYBIATAB0.BIATABID = "CREDOC"
oldYBIATAB0.BIATABK1 = lBIATABK1
oldYBIATAB0.BIATABK2 = ""
oldYBIATAB0.BIATABTXT = ""

lstParam_BIATABK2.Clear
lstParam_BIATABK2.AddItem "Ajouter un enregistrement"

xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CREDOC'" _
     & " and BIATABK1 = '" & lBIATABK1 & "'"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    lstParam_BIATABK2.AddItem rsSab("BIATABK2") & " " & rsSab("BIATABTXT")
    rsSab.MoveNext
Loop

paramCREDOC_Init

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

Private Sub lstMT707_Click()
Dim X As String, K As Integer
On Error GoTo Error_Handler
lstMT707.Visible = False
X = lstMT707.Text
K = InStr(X, "(")
mSWISABSWID_707 = Val(Mid$(X, K + 1, Len(X) - K - 1))

Exit Sub

Error_Handler:
    Call MsgBox(Error, vbCritical, "SAB_DOS_CDO.lstMT707_Click")


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

Private Sub mnuParam_Modèles_CREDOC_Des_Click()
Dim X As String, K As Integer, K1 As Integer
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass
fraParam_Courrier.Visible = False
fgParam_Courrier.Visible = False

If fgParam_Courrier.Rows <= 2 Then
    fgParam_Courrier.Rows = 1
    X = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CREDOC'" _
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
    Call MsgBox(Error, vbCritical, "mnuParam_Modèles_CREDOC_Des_Click")
    Me.Enabled = True: Me.MousePointer = 0
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


Private Sub mnuParam_Modèles_CREDOC_Copier_Click()
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass
msFileSystem.CopyFile paramCDO_Dossier_Path_DROPI & "Modèles\" & lstParam_Modèles_CREDOC.Text, libParam_Modèles_Temp_Path & lstParam_Modèles_CREDOC.Text
lstParam_Modèles_Init

Me.Enabled = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
    Call MsgBox(Error, vbCritical, "mnuParam_Modèles_CREDOC_Copier")
    Me.Enabled = True: Me.MousePointer = 0
    
End Sub

Private Sub mnuParam_Modèles_CREDOC_Delete_Click()
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass
If MsgBox("Confirmez_vous la suppression du modèle : " & vbCrLf & lstParam_Modèles_CREDOC.Text, vbYesNo, "Gestion des modèles CREDOC") = vbYes Then
    msFileSystem.DeleteFile paramCDO_Dossier_Path & "Modèles\" & lstParam_Modèles_CREDOC.Text
    lstParam_Modèles_Init
    
    If blnCourrier_Doc_Exist Then Call sqlYBIATAB0_Transaction("Delete", newCourrier_Doc, oldCourrier_Doc)
    If blnCourrier_Des_Exist Then Call sqlYBIATAB0_Transaction("Delete", newCourrier_Des, oldCourrier_Des)
    oldYBIATAB0 = newCourrier_Des
    Call sqlYBIATAB0_Transaction("Delete_#SAB", newCourrier_Doc, oldCourrier_Doc)

End If
Me.Enabled = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
    Call MsgBox(Error, vbCritical, "mnuParam_Modèles_CREDOC_Delete")
    Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuParam_Modèles_CREDOC_Rename_Click()
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim K As Integer, X As String, xFrom As String, Xto As String
K = InStr(lstParam_Modèles_CREDOC.Text, ".")
If K > 0 Then
    X = Mid$(lstParam_Modèles_CREDOC, 1, K - 1)
Else
    X = lstParam_Modèles_CREDOC
End If
Xto = Trim(InputBox("Nouveau nom du modèle:", "Gestion des modèles", X))
If Xto <> "" Then
    xFrom = paramCDO_Dossier_Path & "Modèles\" & lstParam_Modèles_CREDOC.Text
    msFileSystem.MoveFile xFrom, Replace(xFrom, X, Xto)

    newCourrier_Doc = oldCourrier_Doc
    newCourrier_Doc.BIATABTXT = Replace(oldCourrier_Doc.BIATABTXT, X, Xto)
    Call sqlYBIATAB0_Transaction("Update", newCourrier_Doc, oldCourrier_Doc)
    
    lstParam_Modèles_Init
End If

Me.Enabled = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
    Call MsgBox(Error, vbCritical, "mnuParam_Modèles_CREDOC_Rename")
    Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuParam_Modèles_Temp_Copier_Click()
On Error GoTo Error_Handler
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass
DoEvents
If Dir(paramCDO_Dossier_Path & "Modèles\" & lstParam_Modèles_Temp.Text) <> "" Then Kill paramCDO_Dossier_Path & "Modèles\" & lstParam_Modèles_Temp.Text
msFileSystem.CopyFile libParam_Modèles_Temp_Path & lstParam_Modèles_Temp.Text, paramCDO_Dossier_Path & "Modèles\" & lstParam_Modèles_Temp.Text

X = cmdParam_Courrier_Doc_Exist(lstParam_Modèles_Temp.Text, True)
Call cmdParam_Courrier_Doc_Fields(paramCDO_Dossier_Path & "Modèles\" & lstParam_Modèles_Temp.Text, X)

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
If MsgBox("Confirmez_vous la suppression du modèle : " & vbCrLf & lstParam_Modèles_CREDOC.Text, vbYesNo, "Gestion des modèles temporaires") = vbYes Then
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
            X = "Caractérisques des courriers CDO"
            Call MSflexGrid_Excel("", "CREDOC", X, fgParam_Recap, fgParam_Recap.Cols - 1)
        End If
        If sstabParam.Tab = 2 And optParam_Recap_SAB Then
            X = "Champs #SAB / courriers CDO"
            Call MSflexGrid_Excel("", "CREDOC", X, fgParam_Recap, fgParam_Recap.Cols - 1)
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
            xObjet = "Caractérisques des courriers CDO"
            xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
             & xObjet
    
            Call MSFlexGrid_SendMail(currentSSIWINMAIL, "CREDOC", xObjet, xMesg, fgParam_Recap, fgParam_Recap.Cols - 1)
        End If
        If sstabParam.Tab = 2 And optParam_Recap_SAB Then
            xObjet = "Champs #SAB / courriers CDO"
            xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
             & xObjet
    
            Call MSFlexGrid_SendMail(currentSSIWINMAIL, "CREDOC", xObjet, xMesg, fgParam_Recap, fgParam_Recap.Cols - 1)
        End If

End Select

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub optCourrier_All_Click()
fgZCDOMOD0.Visible = False: lstMT707.Visible = False: lstMT799.Visible = False: lstMT734.Visible = False
lstCourrier_Load

End Sub

Private Sub optCourrier_MOD_Click()
Dim K As Integer
fgSelect.Row = 0
Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
blnZCDOUTI0_Select = False
optCourrier_UTI.Visible = False
        
        fgZCDOMOD0.Row = fgZCDOMOD0.Rows - 1
        Call fgZCDOMOD0_MouseDown(1, 0, 1000, 400)

fgZCDOMOD0.Visible = True

If lstMT707.ListCount > 0 Then

    For K = 0 To lstMT707.ListCount - 1
        If lstMT707.Selected(K) Then lstMT707.Selected(K) = False
    Next K

    lstMT707.Visible = True
End If

If lstMT799.ListCount > 0 Then

    For K = 0 To lstMT799.ListCount - 1
        If lstMT799.Selected(K) Then lstMT799.Selected(K) = False
    Next K

    lstMT799.Visible = True
End If
If lstMT734.ListCount > 0 Then

    For K = 0 To lstMT734.ListCount - 1
        If lstMT734.Selected(K) Then lstMT734.Selected(K) = False
    Next K

    lstMT734.Visible = True
End If

'lstCourrier_Load
End Sub

Private Sub optCourrier_MOD_Validate(Cancel As Boolean)
'fgZCDOMOD0.Visible = False

End Sub

Private Sub optCourrier_OUV_Click()
fgSelect.Row = 0
Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
blnZCDOUTI0_Select = False
optCourrier_UTI.Visible = False
fgZCDOMOD0.Visible = False: lstMT707.Visible = False: lstMT799.Visible = False: lstMT734.Visible = False
lstMT707.Visible = False: lstMT799.Visible = False: lstMT734.Visible = False
blnZCDOMOD0_Select = False
Call fgZCDOMOD0_Color(fgZCDOMOD0_RowClick, MouseMoveUsr.BackColor, fgZCDOMOD0_ColorClick)
fgZCDOMOD0.Visible = False
lstCourrier_Load
End Sub


Private Sub optCourrier_UTI_Click()
fgZCDOMOD0.Visible = False: lstMT707.Visible = False: lstMT799.Visible = False: lstMT734.Visible = False

blnZCDOMOD0_Select = False
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
lstParam_Modèles_CREDOC.Clear
Set objFolder = msFileSystem.GetFolder(paramCDO_Dossier_Path_DROPI & "Modèles")
Set objFiles = objFolder.Files
For Each fsoFile In objFiles
    If InStr(fsoFile.Type, "Document Microsoft Office Word") > 0 Then lstParam_Modèles_CREDOC.AddItem fsoFile.Name
Next
lstParam_Modèles_CREDOC.Visible = True
'___________________________________________________________________________________________________
oldYBIATAB0.BIATABID = "CREDOC"
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
    lstParam_BIATABK1.AddItem "CommissionFR"
    lstParam_BIATABK1.AddItem "CommissionGB"
    
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

Public Sub paramCREDOC_Init()
Dim X As String

X = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CREDOC'" _
     & " and BIATABK1 = 'FAX' and BIATABK2 ='" & usrName_UCase & "'"
Set rsSab = cnsab.Execute(X)
If rsSab.EOF Then
    X = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CREDOC'" _
         & " and BIATABK1 = 'FAX' and BIATABK2 ='*'"
    Set rsSab = cnsab.Execute(X)
End If
If rsSab.EOF Then
    paramCREDOC_FAX = ""
Else
    paramCREDOC_FAX = Trim(rsSab("BIATABTXT"))
End If

X = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CREDOC'" _
     & " and BIATABK1 = 'TEL_NEGO' and BIATABK2 ='" & usrName_UCase & "'"
Set rsSab = cnsab.Execute(X)
If rsSab.EOF Then
    X = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CREDOC'" _
         & " and BIATABK1 = 'TEL_NEGO' and BIATABK2 ='*'"
    Set rsSab = cnsab.Execute(X)
End If
If rsSab.EOF Then
     paramCREDOC_TEL_NEGO = ""
Else
   paramCREDOC_TEL_NEGO = Trim(rsSab("BIATABTXT"))
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


Public Sub cmdPrint_Courrier_Info_M()

Dim X As String, blnInfo_M_Ok As Boolean, K As Integer
On Error GoTo Error_Handler

currentAction = "cmdPrint_Courrier_Info_M"

fgInfo_M.Rows = 1
fgInfo_M.FormatString = fgInfo_M_FormatString
fgInfo_M.Row = 0
blnInfo_M_Ok = True
mUTI_DOC_Index = 0: blnUTI_DOC_Ok = False
mUTI_Com_Index = 0: blnUTI_Com_Ok = False

With appWord.Selection.Find
    .Wrap = wdFindContinue
    .Text = "?NOSTRO_XXX"
    .Replacement.Text = "?NOSTRO_" & xZCDODOS0.CDODOSDEV
    .Execute Replace:=wdReplaceAll
End With




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
         
         Select Case X
           ' Case "?ANNEXES_NB"
           '         fgInfo_M.CellBackColor = mColor_G0
            '        Select Case mANNEXES_NB
            '            Case 0: fgInfo_M.Text = "aucun document annexe"
            '            Case 1: fgInfo_M.Text = "UN document annexe"
            '            Case 2: fgInfo_M.Text = mANNEXES_NB & " documents annexes"
            '        End Select
              Case "?UTI_DOC"
                    mUTI_DOC_Index = K
                    fgInfo_M.Col = 2
                    
                    If blnUTI_DOC_Ok Then
                        fgInfo_M.CellBackColor = mColor_G0
                    Else
                        blnInfo_M_Ok = False
                    End If
               Case "?ESCOMPTE"
                    mUTI_Com_Index = K
                    fgInfo_M.Col = 2
                    
                    If blnUTI_Com_Ok Then
                        fgInfo_M.CellBackColor = mColor_G0
                    Else
                        blnInfo_M_Ok = False
                    End If
             Case Else
                
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
                            Case "?DESCRIPTION":
                                            fgInfo_M.Col = 2
                                            fgInfo_M.CellBackColor = mColor_G0
                                            fgInfo_M.Text = mDescription
                            Case "?UTI_IRR":
                                            fgInfo_M.Col = 2
                                            fgInfo_M.CellBackColor = mColor_G0
                                            fgInfo_M.Text = mIrrégularités
                             Case "?UTI_BEC":
                                            fgInfo_M.Col = 2
                                            fgInfo_M.CellBackColor = mColor_G0
                                            fgInfo_M.Text = mUTI_BEC
                             Case "?UTI_IBAN":
                                            arrFields_BIA_Value(K) = txtRTF_ZCDOSWI0
                                            If arrFields_BIA_Value(K) = "" Then
                                                blnInfo_M_Ok = False
                                            Else
                                                fgInfo_M.Col = 2
                                                fgInfo_M.CellBackColor = mColor_G0
                                                fgInfo_M.Text = arrFields_BIA_Value(K)
                                            End If
                              Case "?AR_ACCORD":
                                            fgInfo_M.Col = 2
                                            fgInfo_M.CellBackColor = mColor_G0
                                            fgInfo_M.Text = mAR_Accord
                              Case "?AR_COURRIER":
                                            fgInfo_M.Col = 2
                                            fgInfo_M.CellBackColor = mColor_G0
                                            fgInfo_M.Text = mAR_Courrier
                               Case "?BQE_RBT":
                                            fgInfo_M.Col = 2
                                            fgInfo_M.CellBackColor = mColor_G0
                                            fgInfo_M.Text = mBQE_RBT
                        Case Else
                                    blnInfo_M_Ok = False
                        End Select
                        
                    End If
            End Select
        'Call lstErr_AddItem(lstErr, cmdContext, "?BIA : " & arrFields_BIA_Name(K)): DoEvents
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
Public Sub cmdPrint_Courrier_Info_M_Replace()
Dim X0 As String, xReplace As String, K As Integer, iLen As Integer, K2 As Integer
On Error GoTo Error_Handler
appWord.Selection.WholeStory
For K = 1 To fgInfo_M.Rows - 1
    fgInfo_M.Row = K
    fgInfo_M.Col = 0: X0 = Trim(fgInfo_M.Text)
        'Call lstErr_AddItem(lstErr, cmdContext, "Replace : " & X): DoEvents

    fgInfo_M.Col = 2: xReplace = Replace(Trim(fgInfo_M.Text), vbCrLf, vbCr)
    iLen = Len(xReplace)
    If iLen < 247 Then
        With appWord.Selection.Find
            '.Wrap = wdFindContinue
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
                '.Wrap = wdFindContinue
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

If blnUTI_Com_Ok Then
    For K = 1 To arrUTI_COM_Escompte_Tbl_Nb
        Call cmdPrint_Courrier_Word_UTI_COM_Escompte(arrUTI_COM_Escompte_Tbl(K))
    Next K
End If
Exit Sub

Error_Handler:

Call MsgBox(Error, vbCritical, currentAction)


End Sub



Public Sub txtRTF_ZCDOTC20()

Dim X As String

mTC2_X = "": mTC2_W = ""

X = "select *  from " & paramIBM_Library_SAB & ".ZCDOTC20 " _
      & " where CDOTC2ETB = " & xZCDODOS0.CDODOSETB & " and CDOTC2AGE = " & xZCDODOS0.CDODOSAGE _
      & " and CDOTC2SER = '" & xZCDODOS0.CDODOSSER & "' and CDOTC2SSE = '" & xZCDODOS0.CDODOSSSE & "'" _
      & " and CDOTC2COP = '" & xZCDODOS0.CDODOSCOP & "' and CDOTC2DOS = " & xZCDODOS0.CDODOSDOS _
      & " and CDOTC2UTI = 0" _
      & " order by CDOTC2SEQ"
      
 Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    If rsSab("CDOTC2MTF") <> 0 Then
        X = Format(rsSab("CDOTC2MTF"), "### ### ##0.00")
    Else
        X = Format(rsSab("CDOTC2TX1"), "##0.00") & " %  "
    End If
    Select Case rsSab("CDOTC2PER")
        Case "T": X = X & " trimestriel"
        Case "U": X = X & " unitaire"
        Case "M": X = X & " mensuel"
        Case "S": X = X & " semestriel"
        Case "A": X = X & " annuel"
        Case Else: X = X & " ?" & rsSab("CDOTC2PER")
    End Select
    
    mTC2_X = mTC2_X & "{\par }" & Trim(rsSab("CDOTC2COM")) & " \tab : " & " " & X _
            & " \highlight2" & rsSab("CDOTC2DEV") & "\highlight0"
    If rsSab("CDOTC2MT1") <> 0 Then mTC2_W = "! barème non affiché"
    rsSab.MoveNext
Loop
'__________________________________________________________________
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#TC2_W", mTC2_W)
mTC2_X = Replace(mTC2_X, "{\par }", "", 1, 1)
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#TC2_X", mTC2_X)

End Sub
Public Function txtRTF_WCDOCOM0_Display() As String

Dim X As String

Select Case WCDOCOM0_X.WCDOCO2PER
    Case "T": X = X & " trimestriel"
    Case "U": X = X & " unitaire"
    Case "M": X = X & " mensuel"
    Case "S": X = X & " semestriel"
    Case "A": X = X & " annuel"
    Case Else: X = X & " ?" ' & rsSab("CDOTC2PER")
End Select

txtRTF_WCDOCOM0_Display = "{\par }" & Trim(WCDOCOM0_X.WCDOCOMCOM) & " \tab : " _
    & Format(WCDOCOM0_X.WCDOCO2TX1, "##0.00") & " %  " _
    & X _
    & "  \highlight2" & WCDOCOM0_X.WCDOCOMDEV & "\highlight0" & "   ht : " & Format(WCDOCOM0_X.WCDOCOMMON, "### ### ##0.00") _
    & " tva : " & Format(WCDOCOM0_X.WCDOCOMMTV, "### ### ##0.00")

End Function

Public Sub txtRTF_ZCDOCOM0_Ouverture()

Dim X As String


Call rsWCDOCOM0_Init(mECNF)
mENOTIF = mECNF
mECSIL = mECNF
mIOUV = mECNF
mCOM_OUV = ""

If xZCDODOS0.CDODOSCON <> "N" Then
    X = "select *  from " & paramIBM_Library_SAB & ".ZCDOCOM0, " & paramIBM_Library_SAB & ".ZCDOCO20 " _
          & " where CDOCOMETB = " & xZCDODOS0.CDODOSETB & " and CDOCOMAGE = " & xZCDODOS0.CDODOSAGE _
          & " and CDOCOMSER = '" & xZCDODOS0.CDODOSSER & "' and CDOCOMSSE = '" & xZCDODOS0.CDODOSSSE & "'" _
          & " and CDOCOMCOP = '" & xZCDODOS0.CDODOSCOP & "' and CDOCOMDOS = " & xZCDODOS0.CDODOSDOS _
          & " and CDOCOMSPE = 1 and CDOCOMCOM in ('ECNF' , 'ECNFPT')" _
          & " and CDOCO2ETB = CDOCOMETB  and CDOCO2AGE = CDOCOMAGE " _
          & " and CDOCO2SER = CDOCOMSER  and CDOCO2SSE = CDOCOMSSE " _
          & " and CDOCO2COP = CDOCOMCOP  and CDOCO2DOS = CDOCOMDOS " _
          & " and CDOCO2NUR = CDOCOMNUR  and CDOCO2UTI = CDOCOMUTI " _
          & " and CDOCO2EVE = CDOCOMEVE  and CDOCO2SEQ = CDOCOMSEQ " _
          & " and CDOCO2SPE = CDOCOMSPE"
          
    Set rsSab = cnsab.Execute(X)
    
    If Not rsSab.EOF Then
        mECNF.WCDOCOMCOM = rsSab("CDOCOMCOM")
        mECNF.WCDOCOMMON = rsSab("CDOCOMMON")
        mECNF.WCDOCOMMTV = rsSab("CDOCOMMTV")
        mECNF.WCDOCOMDEV = rsSab("CDOCOMDEV")
        mECNF.WCDOCO2TX1 = rsSab("CDOCO2TX1")
        mECNF.WCDOCO2PER = rsSab("CDOCO2PER")
        mECNF.WCDOCO2MIN = rsSab("CDOCO2MIN") / 100
        
        WCDOCOM0_X = mECNF
        mCOM_OUV = txtRTF_WCDOCOM0_Display
    End If
    
    X = "select *  from " & paramIBM_Library_SAB & ".ZCDOCOM0, " & paramIBM_Library_SAB & ".ZCDOCO20 " _
          & " where CDOCOMETB = " & xZCDODOS0.CDODOSETB & " and CDOCOMAGE = " & xZCDODOS0.CDODOSAGE _
          & " and CDOCOMSER = '" & xZCDODOS0.CDODOSSER & "' and CDOCOMSSE = '" & xZCDODOS0.CDODOSSSE & "'" _
          & " and CDOCOMCOP = '" & xZCDODOS0.CDODOSCOP & "' and CDOCOMDOS = " & xZCDODOS0.CDODOSDOS _
          & " and CDOCOMSPE = 1 and CDOCOMCOM = 'ECSIL'" _
          & " and CDOCO2ETB = CDOCOMETB  and CDOCO2AGE = CDOCOMAGE " _
          & " and CDOCO2SER = CDOCOMSER  and CDOCO2SSE = CDOCOMSSE " _
          & " and CDOCO2COP = CDOCOMCOP  and CDOCO2DOS = CDOCOMDOS " _
          & " and CDOCO2NUR = CDOCOMNUR  and CDOCO2UTI = CDOCOMUTI " _
          & " and CDOCO2EVE = CDOCOMEVE  and CDOCO2SEQ = CDOCOMSEQ " _
          & " and CDOCO2SPE = CDOCOMSPE "
          
    Set rsSab = cnsab.Execute(X)
    
    If Not rsSab.EOF Then
        mECSIL.WCDOCOMCOM = rsSab("CDOCOMCOM")
        mECSIL.WCDOCOMMON = rsSab("CDOCOMMON")
        mECSIL.WCDOCOMMTV = rsSab("CDOCOMMTV")
        mECSIL.WCDOCOMDEV = rsSab("CDOCOMDEV")
        mECSIL.WCDOCO2TX1 = rsSab("CDOCO2TX1")
        mECSIL.WCDOCO2PER = rsSab("CDOCO2PER")
        mECSIL.WCDOCO2MIN = rsSab("CDOCO2MIN") / 100
        WCDOCOM0_X = mECSIL
        mCOM_OUV = mCOM_OUV & txtRTF_WCDOCOM0_Display
    
    End If
    
End If

'__________________________________________________________________

If xZCDODOS0.CDODOSCON <> "C" Then
    X = "select *  from " & paramIBM_Library_SAB & ".ZCDOCOM0, " & paramIBM_Library_SAB & ".ZCDOCO20 " _
          & " where CDOCOMETB = " & xZCDODOS0.CDODOSETB & " and CDOCOMAGE = " & xZCDODOS0.CDODOSAGE _
          & " and CDOCOMSER = '" & xZCDODOS0.CDODOSSER & "' and CDOCOMSSE = '" & xZCDODOS0.CDODOSSSE & "'" _
          & " and CDOCOMCOP = '" & xZCDODOS0.CDODOSCOP & "' and CDOCOMDOS = " & xZCDODOS0.CDODOSDOS _
          & " and CDOCOMSPE = 999 and CDOCOMCOM = 'ENOTIF'" _
          & " and CDOCO2ETB = CDOCOMETB  and CDOCO2AGE = CDOCOMAGE " _
          & " and CDOCO2SER = CDOCOMSER  and CDOCO2SSE = CDOCOMSSE " _
          & " and CDOCO2COP = CDOCOMCOP  and CDOCO2DOS = CDOCOMDOS " _
          & " and CDOCO2NUR = CDOCOMNUR  and CDOCO2UTI = CDOCOMUTI " _
          & " and CDOCO2EVE = CDOCOMEVE  and CDOCO2SEQ = CDOCOMSEQ " _
          & " and CDOCO2SPE = CDOCOMSPE  order by CDOCOMDBP desc"
          
    Set rsSab = cnsab.Execute(X)
    
    If Not rsSab.EOF Then
        mENOTIF.WCDOCOMCOM = rsSab("CDOCOMCOM")
        mENOTIF.WCDOCOMMON = rsSab("CDOCOMMON")
        mENOTIF.WCDOCOMMTV = rsSab("CDOCOMMTV")
        mENOTIF.WCDOCOMDEV = rsSab("CDOCOMDEV")
        mENOTIF.WCDOCO2TX1 = rsSab("CDOCO2TX1")
        mENOTIF.WCDOCO2PER = rsSab("CDOCO2PER")
        mENOTIF.WCDOCO2MIN = rsSab("CDOCO2MIN") / 100
        WCDOCOM0_X = mENOTIF
        mCOM_OUV = mCOM_OUV & txtRTF_WCDOCOM0_Display
    
    End If
End If
'__________________________________________________________________

If xZCDODOS0.CDODOSCOP = "CDI" Then
    X = "select *  from " & paramIBM_Library_SAB & ".ZCDOCOM0, " & paramIBM_Library_SAB & ".ZCDOCO20 " _
          & " where CDOCOMETB = " & xZCDODOS0.CDODOSETB & " and CDOCOMAGE = " & xZCDODOS0.CDODOSAGE _
          & " and CDOCOMSER = '" & xZCDODOS0.CDODOSSER & "' and CDOCOMSSE = '" & xZCDODOS0.CDODOSSSE & "'" _
          & " and CDOCOMCOP = '" & xZCDODOS0.CDODOSCOP & "' and CDOCOMDOS = " & xZCDODOS0.CDODOSDOS _
          & " and CDOCOMSPE = 1 and CDOCOMCOM = 'IOUV'" _
          & " and CDOCO2ETB = CDOCOMETB  and CDOCO2AGE = CDOCOMAGE " _
          & " and CDOCO2SER = CDOCOMSER  and CDOCO2SSE = CDOCOMSSE " _
          & " and CDOCO2COP = CDOCOMCOP  and CDOCO2DOS = CDOCOMDOS " _
          & " and CDOCO2NUR = CDOCOMNUR  and CDOCO2UTI = CDOCOMUTI " _
          & " and CDOCO2EVE = CDOCOMEVE  and CDOCO2SEQ = CDOCOMSEQ " _
          & " and CDOCO2SPE = CDOCOMSPE "
          
    Set rsSab = cnsab.Execute(X)
    
    If Not rsSab.EOF Then
        mIOUV.WCDOCOMCOM = rsSab("CDOCOMCOM")
        mIOUV.WCDOCOMMON = rsSab("CDOCOMMON")
        mIOUV.WCDOCOMMTV = rsSab("CDOCOMMTV")
        mIOUV.WCDOCOMDEV = rsSab("CDOCOMDEV")
        mIOUV.WCDOCO2TX1 = rsSab("CDOCO2TX1")
        mIOUV.WCDOCO2PER = rsSab("CDOCO2PER")
        mIOUV.WCDOCO2MIN = rsSab("CDOCO2MIN") / 100
        WCDOCOM0_X = mIOUV
        mCOM_OUV = mCOM_OUV & txtRTF_WCDOCOM0_Display
    
    End If
End If

mCOM_OUV = Replace(mCOM_OUV, "{\par }", "", 1, 1)

txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#COM_OUV", mCOM_OUV)

End Sub

Public Sub txtRTF_ZCDOTCO0_CDE()

Dim X As String
On Error GoTo Error_Handler

If Not blnZCDOTCO0_CDE Then
    
    Call rsWCDOCOM0_Init(mELVD)
    mIDOCIR = mELVD
    mEPDIF = mELVD
    mEMODIF = mELVD
    mERFA = mELVD
    mECSIL = mELVD
    If Trim(mECNF.WCDOCOMDEV) <> "" Then
        X = mECNF.WCDOCOMDEV
    Else
        X = mENOTIF.WCDOCOMDEV
    End If
     If Trim(X) = "" Then X = xZCDODOS0.CDODOSDEV
   
    X = "select *  from " & paramIBM_Library_SAB & ".ZCDOTCO0" _
          & " where CDOTCOETA = " & xZCDODOS0.CDODOSETB & " and CDOTCOAGE = 0" _
          & " and CDOTCOCOM in ('ELVD' , 'IDOCIR' , 'EPDIF' , 'EMODIF' , 'ERFA', 'ECSIL') and CDOTCODEV = '" & X & "'"
              
    Set rsSabX = cnsab.Execute(X)
        
    Do Until rsSabX.EOF
        Select Case Trim(rsSabX("CDOTCOCOM"))
            Case "ELVD"
                blnZCDOTCO0_CDE = True
                mELVD.WCDOCOMCOM = rsSabX("CDOTCOCOM")
                mELVD.WCDOCOMMON = rsSabX("CDOTCOMTF")
                mELVD.WCDOCOMDEV = rsSabX("CDOTCODEV")
                mELVD.WCDOCO2TX1 = rsSabX("CDOTCOTX1")
                mELVD.WCDOCO2PER = rsSabX("CDOTCOPER")
                mELVD.WCDOCO2MIN = rsSabX("CDOTCOMIN")
            Case "IDOCIR"
                blnZCDOTCO0_CDE = True
                mIDOCIR.WCDOCOMCOM = rsSabX("CDOTCOCOM")
                mIDOCIR.WCDOCOMMON = rsSabX("CDOTCOMTF")
                mIDOCIR.WCDOCOMDEV = rsSabX("CDOTCODEV")
                mIDOCIR.WCDOCO2TX1 = rsSabX("CDOTCOTX1")
                mIDOCIR.WCDOCO2PER = rsSabX("CDOTCOPER")
                mIDOCIR.WCDOCO2MIN = rsSabX("CDOTCOMIN")
                If mIDOCIR.WCDOCO2MIN = 0 Then mIDOCIR.WCDOCO2MIN = mIDOCIR.WCDOCOMMON
            Case "EPDIF"
                blnZCDOTCO0_CDE = True
                mEPDIF.WCDOCOMCOM = rsSabX("CDOTCOCOM")
                mEPDIF.WCDOCOMMON = rsSabX("CDOTCOMTF")
                mEPDIF.WCDOCOMDEV = rsSabX("CDOTCODEV")
                mEPDIF.WCDOCO2TX1 = rsSabX("CDOTCOTX1")
                mEPDIF.WCDOCO2PER = rsSabX("CDOTCOPER")
                mEPDIF.WCDOCO2MIN = rsSabX("CDOTCOMIN")
                If mEPDIF.WCDOCO2MIN = 0 Then mEPDIF.WCDOCO2MIN = mEPDIF.WCDOCOMMON
            Case "EMODIF"
                blnZCDOTCO0_CDE = True
                mEMODIF.WCDOCOMCOM = rsSabX("CDOTCOCOM")
                mEMODIF.WCDOCOMMON = rsSabX("CDOTCOMTF")
                mEMODIF.WCDOCOMDEV = rsSabX("CDOTCODEV")
                mEMODIF.WCDOCO2TX1 = rsSabX("CDOTCOTX1")
                mEMODIF.WCDOCO2PER = rsSabX("CDOTCOPER")
                mEMODIF.WCDOCO2MIN = rsSabX("CDOTCOMIN")
                If mEMODIF.WCDOCO2MIN = 0 Then mEMODIF.WCDOCO2MIN = mEMODIF.WCDOCOMMON
            Case "ERFA"
                blnZCDOTCO0_CDE = True
                mERFA.WCDOCOMCOM = rsSabX("CDOTCOCOM")
                mERFA.WCDOCOMMON = rsSabX("CDOTCOMTF")
                mERFA.WCDOCOMDEV = rsSabX("CDOTCODEV")
                mERFA.WCDOCO2TX1 = rsSabX("CDOTCOTX1")
                mERFA.WCDOCO2PER = rsSabX("CDOTCOPER")
                mERFA.WCDOCO2MIN = rsSabX("CDOTCOMIN")
                If mERFA.WCDOCO2MIN = 0 Then mERFA.WCDOCO2MIN = mERFA.WCDOCOMMON
             Case "ECSIL"
                blnZCDOTCO0_CDE = True
                mECSIL.WCDOCOMCOM = rsSabX("CDOTCOCOM")
                mECSIL.WCDOCOMMON = rsSabX("CDOTCOMTF")
                mECSIL.WCDOCOMDEV = rsSabX("CDOTCODEV")
                mECSIL.WCDOCO2TX1 = rsSabX("CDOTCOTX1")
                mECSIL.WCDOCO2PER = rsSabX("CDOTCOPER")
                mECSIL.WCDOCO2MIN = rsSabX("CDOTCOMIN")
                If mECSIL.WCDOCO2MIN = 0 Then mECSIL.WCDOCO2MIN = mECSIL.WCDOCOMMON
       End Select
        
     rsSabX.MoveNext
    Loop

End If
GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbCritical, currentAction
Exit_sub:


End Sub
Public Sub txtRTF_ZCDOTCO0_CDI()

Dim X As String
On Error GoTo Error_Handler

If Not blnZCDOTCO0_CDI Then
    
    Call rsWCDOCOM0_Init(mILVD)
   '$JPL 2014-11-05 mIOUV = mILVD
    mIDOCIR = mILVD
    mIPDIF = mILVD
    mIMODIF = mILVD
    mIRFA = mILVD
    mIACD = mILVD
    mIACCEP = mILVD
    mEACCEP = mILVD
    mEACCED = mILVD

'$JPL 2014-11-06_________________________________
'    X = "select *  from " & paramIBM_Library_SAB & ".ZCDOTCO0" _
'          & " where CDOTCOETA = " & xZCDODOS0.CDODOSETB & " and CDOTCOAGE = 0" _
'          & " and CDOTCOCOM in ('IOUV','ILVD','IDOCIR','IPDIF','IMODIF','IRFA','IACD','IACCEP','EACCEP','EACCED')" _
'          & " and CDOTCODEV = '" & xZCDODOS0.CDODOSDEV & "'"
    X = "select *  from " & paramIBM_Library_SAB & ".ZCDOTCO0" _
          & " where CDOTCOETA = " & xZCDODOS0.CDODOSETB & " and CDOTCOAGE = 0" _
          & " and CDOTCOCOM in ('ILVD','IDOCIR','IPDIF','IMODIF','IRFA','IACD','IACCEP','EACCEP','EACCED')" _
          & " and CDOTCODEV = '" & xZCDODOS0.CDODOSDEV & "'"
 '____________________________________________________________________
    Set rsSabX = cnsab.Execute(X)
        
    Do Until rsSabX.EOF
        Select Case Trim(rsSabX("CDOTCOCOM"))
 '$JPL 2014-11-06_________________________________
           'Case "IOUV"
           '     blnZCDOTCO0_CDI = True
           '     mIOUV.WCDOCOMCOM = rsSabX("CDOTCOCOM")
           '     mIOUV.WCDOCOMMON = rsSabX("CDOTCOMTF")
           '     mIOUV.WCDOCOMDEV = rsSabX("CDOTCODEV")
           '     mIOUV.WCDOCO2TX1 = rsSabX("CDOTCOTX1")
           '     mIOUV.WCDOCO2PER = rsSabX("CDOTCOPER")
           '     mIOUV.WCDOCO2MIN = rsSabX("CDOTCOMIN")
  '____________________________________________________________________
           Case "ILVD"
                blnZCDOTCO0_CDI = True
                mILVD.WCDOCOMCOM = rsSabX("CDOTCOCOM")
                mILVD.WCDOCOMMON = rsSabX("CDOTCOMTF")
                mILVD.WCDOCOMDEV = rsSabX("CDOTCODEV")
                mILVD.WCDOCO2TX1 = rsSabX("CDOTCOTX1")
                mILVD.WCDOCO2PER = rsSabX("CDOTCOPER")
                mILVD.WCDOCO2MIN = rsSabX("CDOTCOMIN")
            Case "IDOCIR"
                blnZCDOTCO0_CDI = True
                mIDOCIR.WCDOCOMCOM = rsSabX("CDOTCOCOM")
                mIDOCIR.WCDOCOMMON = rsSabX("CDOTCOMTF")
                mIDOCIR.WCDOCOMDEV = rsSabX("CDOTCODEV")
                mIDOCIR.WCDOCO2TX1 = rsSabX("CDOTCOTX1")
                mIDOCIR.WCDOCO2PER = rsSabX("CDOTCOPER")
                mIDOCIR.WCDOCO2MIN = rsSabX("CDOTCOMIN")
                If mIDOCIR.WCDOCO2MIN = 0 Then mIDOCIR.WCDOCO2MIN = mIDOCIR.WCDOCOMMON
            Case "IPDIF"
                blnZCDOTCO0_CDI = True
                mIPDIF.WCDOCOMCOM = rsSabX("CDOTCOCOM")
                mIPDIF.WCDOCOMMON = rsSabX("CDOTCOMTF")
                mIPDIF.WCDOCOMDEV = rsSabX("CDOTCODEV")
                mIPDIF.WCDOCO2TX1 = rsSabX("CDOTCOTX1")
                mIPDIF.WCDOCO2PER = rsSabX("CDOTCOPER")
                mIPDIF.WCDOCO2MIN = rsSabX("CDOTCOMIN")
                If mIPDIF.WCDOCO2MIN = 0 Then mIPDIF.WCDOCO2MIN = mIPDIF.WCDOCOMMON
            Case "IMODIF"
                blnZCDOTCO0_CDI = True
                mIMODIF.WCDOCOMCOM = rsSabX("CDOTCOCOM")
                mIMODIF.WCDOCOMMON = rsSabX("CDOTCOMTF")
                mIMODIF.WCDOCOMDEV = rsSabX("CDOTCODEV")
                mIMODIF.WCDOCO2TX1 = rsSabX("CDOTCOTX1")
                mIMODIF.WCDOCO2PER = rsSabX("CDOTCOPER")
                mIMODIF.WCDOCO2MIN = rsSabX("CDOTCOMIN")
                If mIMODIF.WCDOCO2MIN = 0 Then mIMODIF.WCDOCO2MIN = mIMODIF.WCDOCOMMON
            Case "IRFA"
                blnZCDOTCO0_CDI = True
                mIRFA.WCDOCOMCOM = rsSabX("CDOTCOCOM")
                mIRFA.WCDOCOMMON = rsSabX("CDOTCOMTF")
                mIRFA.WCDOCOMDEV = rsSabX("CDOTCODEV")
                mIRFA.WCDOCO2TX1 = rsSabX("CDOTCOTX1")
                mIRFA.WCDOCO2PER = rsSabX("CDOTCOPER")
                mIRFA.WCDOCO2MIN = rsSabX("CDOTCOMIN")
                If mIRFA.WCDOCO2MIN = 0 Then mIRFA.WCDOCO2MIN = mIRFA.WCDOCOMMON
              Case "IACD"
                blnZCDOTCO0_CDI = True
                mIACD.WCDOCOMCOM = rsSabX("CDOTCOCOM")
                mIACD.WCDOCOMMON = rsSabX("CDOTCOMTF")
                mIACD.WCDOCOMDEV = rsSabX("CDOTCODEV")
                mIACD.WCDOCO2TX1 = rsSabX("CDOTCOTX1")
                mIACD.WCDOCO2PER = rsSabX("CDOTCOPER")
                mIACD.WCDOCO2MIN = rsSabX("CDOTCOMIN")
            Case "IACCEP"
                blnZCDOTCO0_CDI = True
                mIACCEP.WCDOCOMCOM = rsSabX("CDOTCOCOM")
                mIACCEP.WCDOCOMMON = rsSabX("CDOTCOMTF")
                mIACCEP.WCDOCOMDEV = rsSabX("CDOTCODEV")
                mIACCEP.WCDOCO2TX1 = rsSabX("CDOTCOTX1")
                mIACCEP.WCDOCO2PER = rsSabX("CDOTCOPER")
                mIACCEP.WCDOCO2MIN = rsSabX("CDOTCOMIN")
             Case "EACCEP"
                blnZCDOTCO0_CDI = True
                mEACCEP.WCDOCOMCOM = rsSabX("CDOTCOCOM")
                mEACCEP.WCDOCOMMON = rsSabX("CDOTCOMTF")
                mEACCEP.WCDOCOMDEV = rsSabX("CDOTCODEV")
                mEACCEP.WCDOCO2TX1 = rsSabX("CDOTCOTX1")
                mEACCEP.WCDOCO2PER = rsSabX("CDOTCOPER")
                mEACCEP.WCDOCO2MIN = rsSabX("CDOTCOMIN")
            Case "EACCED"
                blnZCDOTCO0_CDI = True
                mEACCED.WCDOCOMCOM = rsSabX("CDOTCOCOM")
                mEACCED.WCDOCOMMON = rsSabX("CDOTCOMTF")
                mEACCED.WCDOCOMDEV = rsSabX("CDOTCODEV")
                mEACCED.WCDOCO2TX1 = rsSabX("CDOTCOTX1")
                mEACCED.WCDOCO2PER = rsSabX("CDOTCOPER")
                mEACCED.WCDOCO2MIN = rsSabX("CDOTCOMIN")
     End Select
        
     rsSabX.MoveNext
    Loop

End If
GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbCritical, currentAction
Exit_sub:
End Sub


Public Sub txtRTF_ZCDODES0()

Dim X As String
mDescription = ""

X = "select *  from " & paramIBM_Library_SAB & ".ZCDODES0" _
      & " where CDODESETB = " & xZCDODOS0.CDODOSETB & " and CDODESAGE = " & xZCDODOS0.CDODOSAGE _
      & " and CDODESSER = '" & xZCDODOS0.CDODOSSER & "' and CDODESSSE = '" & xZCDODOS0.CDODOSSSE & "'" _
      & " and CDODESCOP = '" & xZCDODOS0.CDODOSCOP & "' and CDODESDOS = " & xZCDODOS0.CDODOSDOS _
      & " order by CDODESNUR , CDODESUTI , CDODESSEQ"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    
    mDescription = mDescription & Trim(rsSab("CDODESTEX")) & "{\par }"
    rsSab.MoveNext
Loop
'__________________________________________________________________
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#DESCRIPTION", mDescription)
mDescription = Replace(mDescription, "{\par }", vbCrLf)
End Sub



Public Sub txtRTF_ZCDOIRR0()

Dim X As String
mIrrégularités = ""
arrFields_BIA_Value(mIrrégularités_Index) = ""

X = "select *  from " & paramIBM_Library_SAB & ".ZCDOIRR0" _
      & " where CDOIRRETB = " & xZCDOUTI0.CDOUTIETB & " and CDOIRRAGE = " & xZCDOUTI0.CDOUTIAGE _
      & " and CDOIRRSER = '" & xZCDOUTI0.CDOUTISER & "' and CDOIRRSSE = '" & xZCDOUTI0.CDOUTISSE & "'" _
      & " and CDOIRRCOP = '" & xZCDOUTI0.CDOUTICOP & "' and CDOIRRDOS = " & xZCDOUTI0.CDOUTIDOS _
      & " and CDOIRRNUR = '" & xZCDOUTI0.CDOUTINUR & "' and CDOIRRUTI = " & xZCDOUTI0.CDOUTIUTI _
      & " order by CDOIRRSEQ"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    
    mIrrégularités = mIrrégularités & Trim(rsSab("CDOIRRTEX")) & vbCrLf
    rsSab.MoveNext
Loop
'__________________________________________________________________
End Sub




Public Function txtRTF_ZCDOSWI0()

Dim X As String
txtRTF_ZCDOSWI0 = ""

If blnZCDOUTI0_Select Then

    X = "select *  from " & paramIBM_Library_SAB & ".ZCDOSWI0" _
          & " where CDOSWIETB = " & xZCDOUTI0.CDOUTIETB & " and CDOSWIAGE = " & xZCDOUTI0.CDOUTIAGE _
          & " and CDOSWISER = '" & xZCDOUTI0.CDOUTISER & "' and CDOSWISSE = '" & xZCDOUTI0.CDOUTISSE & "'" _
          & " and CDOSWICOP = '" & xZCDOUTI0.CDOUTICOP & "' and CDOSWIDOS = " & xZCDOUTI0.CDOUTIDOS _
          & " and CDOSWINUR = '" & xZCDOUTI0.CDOUTINUR & "' and CDOSWIUTI = " & xZCDOUTI0.CDOUTIUTI _
          & " and CDOSWIBEN = '" & xZCDODOS0.CDODOSBEN & "' and CDOSWIBER = '" & xZCDODOS0.CDODOSBER & "'" _
          & " order by CDOSWIPAI , CDOSWIREG"
    Set rsSabX = cnsab.Execute(X)
    
    Do While Not rsSabX.EOF
                X = Trim(rsSabX("CDOSWIIBE"))
                If X <> "" Then txtRTF_ZCDOSWI0 = Format$(X, "&&&& &&&& &&&& &&&& &&&& &&&& &&&&!")
        
        rsSabX.MoveNext
    Loop
End If
'__________________________________________________________________
End Function

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

xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CREDOC'" _
     & " and BIATABK1 = 'Courrier_Doc' and BIATABTXT = '" & Trim(lFileName) & "'"
Set rsSab_Local = cnsab.Execute(xSql)
If Not rsSab_Local.EOF Then
    cmdParam_Courrier_Doc_Exist = rsSab_Local("BIATABK2")
Else
    If blnAdd Then
    
        Dim oldYBIATAB0_Local As typeYBIATAB0, newYBIATAB0_Local As typeYBIATAB0
        
        X = "select BIATABK2 from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CREDOC'" _
             & " and BIATABK1 = 'Courrier_Doc' order by BIATABK2 desc"
        Set rsSab_Local = cnsab.Execute(X)
        oldYBIATAB0_Local.BIATABID = "CREDOC"
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

newYBIATAB0_Local.BIATABID = "CREDOC_#SAB"
newYBIATAB0_Local.BIATABK1 = lBIATABK1
newYBIATAB0_Local.BIATABTXT = ""



'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
xSql = "Delete  from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CREDOC_#SAB'" _
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



Public Sub txtRTF_YSWISAB0()
Dim X As String, xSql As String
Dim xValue As String, xRTF As String, V
Dim K As Integer, K2 As Integer, iAsc13 As Integer, iLen As Integer
Dim xField As String
On Error GoTo Error_Handler

If Not blnSIDE_DB_Open Then
    cnSIDE_DB.Open paramODBC_DSN_SIDE_DB
    blnSIDE_DB_Open = True
End If

txtRTF_MT700 = ""
mBQE_RBT = ""
mSWISABSWID_707 = 0: lstMT707.Clear
mSWISABSWID_799 = 0: lstMT799.Clear
mSWISABSWID_734 = 0: lstMT734.Clear
'__________________________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABOPEC = '" & xZCDODOS0.CDODOSCOP & "'" _
     & " and   SWISABOPEN = " & xZCDODOS0.CDODOSDOS _
     & " and  SWISABWMTK in('700' , '701' , '707','799','734') and SWISABWES = 'E' order by SWISABSWID"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF

    
    If rsSab("SWISABWMTK") = "707" Then
        lstMT707.AddItem rsSab("SWISABWMTK") _
                        & " reçu le " & dateImp10_S(rsSab("SWISABWAMJ")) & " " & timeImp8(rsSab("SWISABWHMS")) _
                        & "  (" & Trim(rsSab("SWISABSWID")) & ")"
        mSWISABSWID_707 = rsSab("SWISABSWID")
    End If
    If rsSab("SWISABWMTK") = "799" Then
        lstMT799.AddItem rsSab("SWISABWMTK") _
                        & " reçu le " & dateImp10_S(rsSab("SWISABWAMJ")) & " " & timeImp8(rsSab("SWISABWHMS")) _
                        & "  (" & Trim(rsSab("SWISABSWID")) & ")"
        mSWISABSWID_799 = rsSab("SWISABSWID")
    End If
    If rsSab("SWISABWMTK") = "734" Then
        lstMT734.AddItem rsSab("SWISABWMTK") _
                        & " reçu le " & dateImp10_S(rsSab("SWISABWAMJ")) & " " & timeImp8(rsSab("SWISABWHMS")) _
                        & "  (" & Trim(rsSab("SWISABSWID")) & ")"
        mSWISABSWID_734 = rsSab("SWISABSWID")
    End If
    
    xRTF = "MT" & rsSab("SWISABWMTK") & " de " & rsSab("SWISABWBIC") _
         & " reçu le " & dateImp10_S(rsSab("SWISABWAMJ")) & " " & timeImp8(rsSab("SWISABWHMS")) _
         & "    (" & Trim(rsSab("SWISABSWID")) & ")"
    txtRTF_MT700 = txtRTF_MT700 & "\par\cf0\highlight2 " & xRTF & "\tab\highlight0\par"
    
    '\cf4\highlight5\i MT700\tab\tab\tab\tab\tab\tab\tab\highlight0\par
    '\cf3\i0 Message\par
    
    'Call arrMT_Type_Scan(rsSab("SWISABWMTK"))
    Call arrMT_Fields_Load(rsSab("SWISABWMTK"))

    xSql = "select *  from rtextField  " _
        & "where Aid = " & rsSab("SWISABWID1") _
        & " and text_s_umidl = " & rsSab("SWISABWIDL") _
        & " and text_s_umidh  =  " & rsSab("SWISABWIDH") _
        & " order by field_cnt"
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    
    xRTF = ""
    
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
            
            If rsSab("SWISABWMTK") = "700" And rsSIDE_DB("field_code") = "53" Then mBQE_RBT = xValue
            
            xField = rsSIDE_DB("field_code") & rsSIDE_DB("field_option")
            xRTF = xRTF & "\cf0\ul\i\par " & xField & " " & arrMT_Fields_Scan(xField) & "\i0\ulnone\cf3 "
           
                iLen = Len(xValue)
                K = 1
                Do
                   iAsc13 = InStr(K, xValue, Asc13)
                   If iAsc13 > 0 Then
                        xRTF = xRTF & "\par " & Trim(Mid$(xValue, K, iAsc13 - K))
                       K = iAsc13 + 2
                   End If
                Loop Until iAsc13 = 0
                
                xRTF = xRTF & "\par " & Trim(Mid$(xValue, K, iLen - K + 1))
            
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
            'X = X & "\cf3 " & xrText.text_data_block
            
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
                        xRTF = xRTF & "\par " & Trim(Mid$(xValue, K, iAsc13 - K))
                    Else
                        K2 = InStr(2, X, ":")
                        If K2 > 0 Then
                            xField = Mid$(X, 2, K2 - 2)
                            'xRTF = xRTF & "\cf0\highlight8\par " & Trim(Mid$(X, 2, K2 - 1)) & " " & arrMT_Fields_Scan(xField) & "\tab\tab\tab\highlight0\cf3\par "
                            xRTF = xRTF & "\cf0\ul\i\par " & Trim(Mid$(X, 2, K2 - 1)) & " " & arrMT_Fields_Scan(xField) & "\i0\ulnone\cf3\par "
                            X = Trim(Mid$(X, K2 + 1, Len(X) - K2))
                             xRTF = xRTF & X
                                Select Case xField
                                    Case "52A", "57A", "57D", "58A", "59A", "59F": xRTF = xRTF & txtRTF_ZSWIBIC0(X)
                                End Select
                        Else
                            xRTF = xRTF & "\par " & Trim(Mid$(xValue, K, iAsc13 - K))
                            
                            If rsSab("SWISABWMTK") = "700" And Mid$(xField, 1, 2) = "53" Then mBQE_RBT = Trim(Mid$(xValue, K, iAsc13 - K))

                        End If
                    End If
                    
                    K = iAsc13 + 2
                End If
             Loop Until iAsc13 = 0
        End If
    End If
'End If

    txtRTF_MT700 = txtRTF_MT700 & xRTF & "\par"
    rsSab.MoveNext

Loop

If lstMT707.ListCount > 0 Then
    lstMT707.Selected(lstMT707.ListCount - 1) = True
    lstMT707.Selected(lstMT707.ListCount - 1) = False
End If
If lstMT799.ListCount > 0 Then
    lstMT799.Selected(lstMT799.ListCount - 1) = True
    lstMT799.Selected(lstMT799.ListCount - 1) = False
End If
If lstMT734.ListCount > 0 Then
    lstMT734.Selected(lstMT734.ListCount - 1) = True
    lstMT734.Selected(lstMT734.ListCount - 1) = False
End If
'_______________________________________________________________________
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "#MT700", txtRTF_MT700)

GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbCritical, currentAction
Exit_sub:

End Sub
Public Function txtRTF_ZSWIBIC0(lBIC As String) As String
Dim xSql As String
xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIBIC0 where SWIBICBIC = '" & Replace(lBIC, "'", "") & "'"
Set rsSabX = cnsab.Execute(xSql)
    
If Not rsSabX.EOF Then
    txtRTF_ZSWIBIC0 = "\par\cf9\tab " & rsSabX("SWIBICIN1") & "\par\tab " & rsSabX("SWIBICVIL")
Else
    txtRTF_ZSWIBIC0 = ""
End If


End Function



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

Public Sub fgUTI_Com_Load()
Dim X As String, xSql As String, K As Integer
Dim rsSABY As New ADODB.Recordset
Dim curX As Currency, curTVA As Currency


On Error GoTo Error_Handler
currentAction = "fgUTI_Com_Load"
fgUTI_Com.Rows = 1
fgUTI_Com.Rows = 21
fgUTI_Com.Row = 0
mUTI_Com_RowMin = 0

'X = "select distinct(CDOCOMREG)  from " & paramIBM_Library_SAB & ".ZCDOCOM0" _
'      & " where CDOCOMETB = " & xZCDODOS0.CDODOSETB & " and CDOCOMAGE = " & xZCDODOS0.CDODOSAGE _
'      & " and CDOCOMSER = '" & xZCDODOS0.CDODOSSER & "' and CDOCOMSSE = '" & xZCDODOS0.CDODOSSSE & "'" _
'      & " and CDOCOMCOP = '" & xZCDODOS0.CDODOSCOP & "' and CDOCOMDOS = " & xZCDODOS0.CDODOSDOS _
 '     & " and CDOCOMNUR = " & xZCDOUTI0.CDOUTINUR & " and CDOCOMUTI = " & xZCDOUTI0.CDOUTIUTI _
'      & " and CDOCOMUTR = 0 and CDOCOMNRE = 0 order by cdocomreg desc"

'$JPL 2013-01-29 CDE 101378 uti 4 + escompte

X = "select distinct(CDOCOMREG)  from " & paramIBM_Library_SAB & ".ZCDOCOM0" _
      & " where CDOCOMETB = " & xZCDODOS0.CDODOSETB & " and CDOCOMAGE = " & xZCDODOS0.CDODOSAGE _
      & " and CDOCOMSER = '" & xZCDODOS0.CDODOSSER & "' and CDOCOMSSE = '" & xZCDODOS0.CDODOSSSE & "'" _
      & " and CDOCOMCOP = '" & xZCDODOS0.CDODOSCOP & "' and CDOCOMDOS = " & xZCDODOS0.CDODOSDOS _
      & " and CDOCOMNUR = " & xZCDOUTI0.CDOUTINUR & " and CDOCOMUTI = " & xZCDOUTI0.CDOUTIUTI _
      & " order by cdocomreg desc"
      
Set rsSabX = cnsab.Execute(X)

If rsSabX.EOF Then
    Call MsgBox("Impossible de déterminer la date de réglement pour cette utilisation", vbCritical, "fgUTI_Com_Load")
    Exit Sub
End If

'X = "select *  from " & paramIBM_Library_SAB & ".ZCDOCOM0" _
'      & " where CDOCOMETB = " & xZCDODOS0.CDODOSETB & " and CDOCOMAGE = " & xZCDODOS0.CDODOSAGE _
'      & " and CDOCOMSER = '" & xZCDODOS0.CDODOSSER & "' and CDOCOMSSE = '" & xZCDODOS0.CDODOSSSE & "'" _
'      & " and CDOCOMCOP = '" & xZCDODOS0.CDODOSCOP & "' and CDOCOMDOS = " & xZCDODOS0.CDODOSDOS _
'      & " and CDOCOMREG = " & rsSabX(0) & " and CDOCOMUTR = 0 and CDOCOMNRE = 0 " _
'      & " order by CDOCOMEVE , CDOCOMSEQ , CDOCOMSPE"
'$JPL 2013-01-29 CDE 101378 uti 4 + escompte

X = "select *  from " & paramIBM_Library_SAB & ".ZCDOCOM0" _
      & " where CDOCOMETB = " & xZCDODOS0.CDODOSETB & " and CDOCOMAGE = " & xZCDODOS0.CDODOSAGE _
      & " and CDOCOMSER = '" & xZCDODOS0.CDODOSSER & "' and CDOCOMSSE = '" & xZCDODOS0.CDODOSSSE & "'" _
      & " and CDOCOMCOP = '" & xZCDODOS0.CDODOSCOP & "' and CDOCOMDOS = " & xZCDODOS0.CDODOSDOS _
      & " and CDOCOMREG = " & rsSabX(0) _
      & " order by CDOCOMEVE , CDOCOMSEQ , CDOCOMSPE"

Set rsSabX = cnsab.Execute(X)

Do Until rsSabX.EOF
    curX = rsSabX("CDOCOMMON")
    If curX > 0 Then
        fgUTI_Com.Row = fgUTI_Com.Row + 1
        
        fgUTI_Com.Col = 1: fgUTI_Com.CellBackColor = mColor_Y1
        fgUTI_Com.Text = Format(curX, "### ### ##0.00")
        
        
        fgUTI_Com.Col = 0: fgUTI_Com.CellBackColor = mColor_Y1
        
        If optLangue_FR Then
                X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
                      & " where BIATABID = 'CREDOC' and BIATABK1 = 'CommissionFR' and BIATABK2 = '" & rsSabX("CDOCOMCOM") & "'"
        Else
                X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
                      & " where BIATABID = 'CREDOC' and BIATABK1 = 'CommissionGB' and BIATABK2 = '" & rsSabX("CDOCOMCOM") & "'"
        End If
        Set rsSABY = cnsab.Execute(X)
        
        If Not rsSABY.EOF Then
            fgUTI_Com.Text = Trim(rsSABY("BIATABTXT"))
        Else
            fgUTI_Com.Text = rsSabX("CDOCOMCOM")
        End If
    
        
        fgUTI_Com.Col = 2: fgUTI_Com.CellBackColor = mColor_Y1
        curTVA = rsSabX("CDOCOMMTV")
        If curTVA <> 0 Then
        
            X = "select *  from " & paramIBM_Library_SAB & ".ZCDOCO20" _
                & " where CDOCO2ETB = " & xZCDODOS0.CDODOSETB & " and CDOCO2AGE = " & xZCDODOS0.CDODOSAGE _
                & " and CDOCO2SER = '" & xZCDODOS0.CDODOSSER & "' and CDOCO2SSE = '" & xZCDODOS0.CDODOSSSE & "'" _
                & " and CDOCO2COP = '" & xZCDODOS0.CDODOSCOP & "' and CDOCO2DOS = " & xZCDODOS0.CDODOSDOS _
                & " and CDOCO2NUR = '" & xZCDOUTI0.CDOUTINUR & "' and CDOCO2UTI = " & xZCDOUTI0.CDOUTIUTI _
                & " and CDOCO2EVE = '" & rsSabX("CDOCOMEVE") & "'" _
                & " and CDOCO2SEQ = " & rsSabX("CDOCOMSEQ") _
                & " and CDOCO2SPE = " & rsSabX("CDOCOMSPE")
            Set rsSABY = cnsab.Execute(X)
            
            If Not rsSABY.EOF Then
                If rsSABY("CDOCO2TVA") <> "O" Then curTVA = 0
            End If
        End If
        If curTVA = 0 Then
            fgUTI_Com.Text = ""
        Else
            fgUTI_Com.Text = Format(curTVA, "### ### ##0.00")
            curX = curX + curTVA
       End If
       
        fgUTI_Com.Col = 3: fgUTI_Com.CellBackColor = mColor_Y1
        fgUTI_Com.Text = Format(curX, "### ### ##0.00")
    
              
            
    End If
    rsSabX.MoveNext
Loop

mUTI_Com_RowMin = fgUTI_Com.Row
fgUTI_Com.Col = 0

If optLangue_FR Then
    fgUTI_Com.Row = fgUTI_Com.Row + 1
    fgUTI_Com.Text = "frais de la banque de remboursement"
    fgUTI_Com.Row = 10: fgUTI_Com.Text = "Total des commissions L/C"
    fgUTI_Com.Row = 11: fgUTI_Com.Text = "Montant net L/C"
    fgUTI_Com.Row = 12: fgUTI_Com.Text = "intérêts"
    fgUTI_Com.Row = 13: fgUTI_Com.Text = "frais de dossier"
    
    fgUTI_Com.Row = 20: fgUTI_Com.Text = "Montant net payé"
Else
    fgUTI_Com.Row = fgUTI_Com.Row + 1
    fgUTI_Com.Text = "frais de la banque de remboursement"
    fgUTI_Com.Row = 10: fgUTI_Com.Text = "documentary credit commissions"
    fgUTI_Com.Row = 11: fgUTI_Com.Text = "net amount on documentary credit"
    fgUTI_Com.Row = 12: fgUTI_Com.Text = "interests"
    fgUTI_Com.Row = 13: fgUTI_Com.Text = "frais de dossier"
    
    fgUTI_Com.Row = 20: fgUTI_Com.Text = "NET PAID"
End If

    fgUTI_Com.Row = 10
    For K = 0 To 3
        fgUTI_Com.Col = K: fgUTI_Com.CellBackColor = mColor_Y2
        fgUTI_Com.CellFontBold = True
    Next K
    fgUTI_Com.Row = 11
    For K = 0 To 3
        fgUTI_Com.Col = K: fgUTI_Com.CellBackColor = &H80C0FF
        fgUTI_Com.CellFontBold = True
        fgUTI_Com.CellForeColor = vbBlue
    Next K
    
    fgUTI_Com.Row = 20
    For K = 0 To 3
        fgUTI_Com.Col = K: fgUTI_Com.CellBackColor = &H80C0FF
        fgUTI_Com.CellFontBold = True
        fgUTI_Com.CellForeColor = vbBlue
    Next K

Call fgUTI_Com_Total

blnUTI_Com_Loaded = True
Exit Sub
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub txtUTI_Com_M_KeyPress(KeyAscii As Integer)
Select Case mUTI_Com_Col
    Case 0
    Case 1, 2: Call num_KeyAsciiD(KeyAscii, txtUTI_Com_M)
    Case Else: KeyAscii = 0
End Select

End Sub


Private Sub txtUTI_Com_M_Validate(Cancel As Boolean)
Dim blnOk As Boolean, K As Integer, curMTD As Currency, curTVA As Currency
Dim retval As Integer

'retval = GetKeyState(vbKeyTab)
'If retval <= 0 Then Cancel = True


fgUTI_Com.Col = 0: fgUTI_Com.CellBackColor = fgUTI_Com.BackColor
txtUTI_Com_M.Visible = False
fgUTI_Com.Col = mUTI_Com_Col

Select Case mUTI_Com_Col
    Case 0: fgUTI_Com.Text = Trim(txtUTI_Com_M)
    Case 1, 2: curMTD = num_CDec(txtUTI_Com_M)
            If curMTD = 0 Then
                fgUTI_Com.Text = ""
            Else
                fgUTI_Com.Text = Format(curMTD, "### ### ##0.00")
            End If
End Select

blnOk = True
fgUTI_Com.Col = 0: If Trim(fgUTI_Com.Text) = "" Then blnOk = False
fgUTI_Com.Col = 1: curMTD = num_CDec(fgUTI_Com.Text)
fgUTI_Com.Col = 2: curTVA = num_CDec(fgUTI_Com.Text)

If Abs(curTVA) > Abs(curMTD) Then blnOk = False: Call MsgBox("montant HT < montant TVA", vbCritical, "Tableau Escompte")

curMTD = curMTD + curTVA
fgUTI_Com.Col = 3
If curMTD = 0 Then
    blnOk = False
    fgUTI_Com.Text = ""
Else
    fgUTI_Com.Text = Format(curMTD, "### ### ##0.00")
End If

Select Case curMTD
    Case 0: blnOk = False
            If curTVA <> 0 Then Call MsgBox("Préciser le montant HT", vbExclamation, "Tableau Escompte")
    Case Is < 0
            If curTVA > 0 Then blnOk = False: Call MsgBox("Vérifier le sens des montants HT / TVA", vbCritical, "Tableau Escompte")
    Case Is > 0
            If curTVA < 0 Then blnOk = False: Call MsgBox("Vérifier le sens des montants HT / TVA", vbCritical, "Tableau Escompte")
End Select


If blnOk Then
    fgUTI_Com.Col = 0: fgUTI_Com.CellBackColor = mColor_G1
    fgUTI_Com.Col = 1: fgUTI_Com.CellBackColor = mColor_G1
    fgUTI_Com.Col = 2: fgUTI_Com.CellBackColor = mColor_G1
    fgUTI_Com.Col = 3: fgUTI_Com.CellBackColor = mColor_G1
Else
    fgUTI_Com.Col = 0: fgUTI_Com.CellBackColor = fgUTI_Com.BackColor
    fgUTI_Com.Col = 1: fgUTI_Com.CellBackColor = fgUTI_Com.BackColor
    fgUTI_Com.Col = 2: fgUTI_Com.CellBackColor = fgUTI_Com.BackColor
    fgUTI_Com.Col = 3: fgUTI_Com.CellBackColor = fgUTI_Com.BackColor

End If

Call fgUTI_Com_Total

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
    fgUTI_DOC_Sort
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
'fgUTI_DOC.Redraw = True
'On Error Resume Next
'If fgUTI_DOC.Visible Then fgUTI_DOC.SetFocus
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

Public Function cmdPrint_Courrier_Word_UTI_IBAN()

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


Public Sub JPL()
Dim X As String

Me.Enabled = False: Me.MousePointer = vbHourglass
oldYBIATAB0.BIATABID = "CREDOC"
oldYBIATAB0.BIATABK1 = "#SAB"
newYBIATAB0 = oldYBIATAB0
Open "C:\Temp\SAB_DOS_CDO\commissions CDI.txt" For Input As #1

Do Until EOF(1)
    Line Input #1, X
    
    If Trim(X) <> "" Then
        newYBIATAB0.BIATABK2 = Trim(Mid$(X, 1, 12))
    
        
        newYBIATAB0.BIATABTXT = Trim(Mid$(X, 13, Len(X) - 12))
        Call sqlYBIATAB0_Transaction("New", newYBIATAB0, oldYBIATAB0)
    End If
Loop
Close 1
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub txtRTF_ZCDODOS0_ADR(lX1 As String, lX2 As String)
Dim X As String

If w_ZADRESSE0.ADRESSTYP = "1" Then
    X = Val(w_ZADRESSE0.ADRESSNUM) & " - "
Else
    X = w_ZADRESSE0.ADRESSTYP & " " & Val(w_ZADRESSE0.ADRESSNUM) & " - "
End If
txtRTF.TextRTF = Replace(txtRTF.TextRTF, lX1, X & Trim(w_ZADRESSE0.ADRESSRA1) & Trim(w_ZADRESSE0.ADRESSRA2))

X = ""
If Trim(w_ZADRESSE0.ADRESSAD1) <> "" Then X = X & Trim(w_ZADRESSE0.ADRESSAD1) & "{\par }"
If Trim(w_ZADRESSE0.ADRESSAD2) <> "" Then X = X & Trim(w_ZADRESSE0.ADRESSAD2) & "{\par }"
If Trim(w_ZADRESSE0.ADRESSAD3) <> "" Then X = X & Trim(w_ZADRESSE0.ADRESSAD3) & "{\par }"
If Trim(w_ZADRESSE0.ADRESSCOP) <> "" Then X = X & Trim(w_ZADRESSE0.ADRESSCOP) & "  "
If Trim(w_ZADRESSE0.ADRESSVIL) <> "" Then X = X & Trim(w_ZADRESSE0.ADRESSVIL) & "  "
If Trim(w_ZADRESSE0.ADRESSPAY) <> "" Then X = X & Trim(w_ZADRESSE0.ADRESSPAY)
If Trim(w_ZADRESSE0.ADRESSTEL) <> "" Then X = X & "{\par }" & "Tél :  " & Trim(w_ZADRESSE0.ADRESSTEL)
If Trim(w_ZADRESSE0.ADRESSFAX) <> "" Then X = X & " - Fax :  " & Trim(w_ZADRESSE0.ADRESSFAX)
If Trim(w_ZADRESSE0.ADRESSTEX) <> "" Then X = X & " - Tlx :  " & Trim(w_ZADRESSE0.ADRESSTEX)

If lX1 = "#BEN_RS" Then
    If mBEN_TVANIFCLIT <> "" Then
        X = X & "{\par\cf0\highlight8 " & "NIF : " & mBEN_TVANIFCLIT & " \highlight0}"
    Else
         X = X & "{\par\cf2\highlight1 NIF : ???? \cf0\highlight0}"
   End If
     If mBEN_CDOTIESRN <> "" Then
        X = X & "{\par\cf0\highlight8 " & "SIREN : " & mBEN_CDOTIESRN & " \highlight0}"
    Else
         X = X & "{\par\cf0 SIREN :  \cf0}"
   End If
  
   '
End If

txtRTF.TextRTF = Replace(txtRTF.TextRTF, lX2, X)
End Sub


Public Sub fgUTI_Com_Total()
Dim K As Integer, curMTD As Currency
Dim curHT_S As Currency, curTVA_S As Currency, curTTC_S As Currency
Dim curHT_T As Currency, curTVA_T As Currency, curTTC_T As Currency

For K = 1 To 9
    fgUTI_Com.Row = K
    fgUTI_Com.Col = 1: curHT_S = curHT_S + num_CDec(fgUTI_Com.Text)
    fgUTI_Com.Col = 2: curTVA_S = curTVA_S + num_CDec(fgUTI_Com.Text)
    fgUTI_Com.Col = 3: curTTC_S = curTTC_S + num_CDec(fgUTI_Com.Text)
    
Next K

fgUTI_Com.Row = 10
fgUTI_Com.Col = 1: fgUTI_Com.Text = Format(curHT_S, "### ### ##0.00")
fgUTI_Com.Col = 2: fgUTI_Com.Text = Format(curTVA_S, "### ### ##0.00")
fgUTI_Com.Col = 3: fgUTI_Com.Text = Format(curTTC_S, "### ### ##0.00")

curUTI_COM_CR = curTTC_S

curMTD = xZCDOUTI0.CDOUTIMPA - curTTC_S
fgUTI_Com.Row = 11
fgUTI_Com.Col = 3: fgUTI_Com.Text = Format(curMTD, "### ### ##0.00")

For K = 12 To 19
    fgUTI_Com.Row = K
    fgUTI_Com.Col = 3: curMTD = curMTD - num_CDec(fgUTI_Com.Text)
    
Next K
fgUTI_Com.Row = 20
fgUTI_Com.Col = 3: fgUTI_Com.Text = Format(curMTD, "### ### ##0.00")
curUTI_NET_Escompte = curMTD
End Sub

Public Function cmdPrint_Courrier_Word_UTI_BLOCAGE_Intitulé() As String
Dim rsSABY As New ADODB.Recordset

cmdPrint_Courrier_Word_UTI_BLOCAGE_Intitulé = ""

X = "select *  from " & paramIBM_Library_SAB & ".ZCDOSWI0" _
      & " where CDOSWIETB = " & rsSabX("CDOREGETB") & " and CDOSWIAGE = " & rsSabX("CDOREGAGE") _
      & " and CDOSWISER = '" & rsSabX("CDOREGSER") & "' and CDOSWISSE = '" & rsSabX("CDOREGSSE") & "'" _
      & " and CDOSWICOP = '" & rsSabX("CDOREGCOP") & "' and CDOSWIDOS = " & rsSabX("CDOREGDOS") _
      & " and CDOSWINUR = " & rsSabX("CDOREGNUR") & " and CDOSWIUTI = " & rsSabX("CDOREGUTI") _
      & " and CDOSWIPAI = " & rsSabX("CDOREGPAI") _
      & " and CDOSWIREG = " & rsSabX("CDOREGREG")
Set rsSABY = cnsab.Execute(X)

If Not rsSABY.EOF Then
        Select Case rsSABY("CDOSWIBER")
            Case " "
                     X = "select CLIENARA1  from " & paramIBM_Library_SAB & ".ZCLIENA0" _
                          & " where CLIENAETB = " & rsSABY("CDOSWIETB") & " and CLIENACLI = " & rsSABY("CDOSWIBEN")
                    Set rsSABY = cnsab.Execute(X)
                    If Not rsSABY.EOF Then cmdPrint_Courrier_Word_UTI_BLOCAGE_Intitulé = Trim(rsSABY("CLIENARA1"))
            Case "B": cmdPrint_Courrier_Word_UTI_BLOCAGE_Intitulé = rsSABY("CDOSWIBBE")
            Case "T"
                    X = "select CDOTIERA1  from " & paramIBM_Library_SAB & ".ZCDOTIE0" _
                          & " where CDOTIEETB = " & rsSABY("CDOSWIETB") & " and CDOTIETIE = '" & rsSABY("CDOSWIBEN") & "'"
                    Set rsSABY = cnsab.Execute(X)
                    If Not rsSABY.EOF Then cmdPrint_Courrier_Word_UTI_BLOCAGE_Intitulé = Trim(rsSABY("CDOTIERA1"))
    
        End Select
    
End If

End Function
