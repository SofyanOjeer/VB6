VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDROPI 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Dossier Risque OPérationnel Informatisé"
   ClientHeight    =   10605
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   14970
   Icon            =   "DROPI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10605
   ScaleWidth      =   14970
   Begin VB.ListBox lstErr 
      BackColor       =   &H00FFFFFA&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7080
      TabIndex        =   16
      Top             =   45
      Width           =   6972
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9975
      Left            =   30
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   17595
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   16777152
      ForeColor       =   8388736
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Risques Opérationnels"
      TabPicture(0)   =   "DROPI.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Paramétrage"
      TabPicture(1)   =   "DROPI.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraParam"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Habilitations"
      TabPicture(2)   =   "DROPI.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraAut"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "."
      TabPicture(3)   =   "DROPI.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraDossier"
      Tab(3).Control(1)=   "fraExport"
      Tab(3).Control(2)=   "lstW"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "DROPI.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame1"
      Tab(4).Control(1)=   "X_cmdUpdate_Dossier"
      Tab(4).Control(2)=   "X_cmdUpdate"
      Tab(4).Control(3)=   "libDossier_ROPINFGTXT"
      Tab(4).Control(4)=   "fraUpdate_ROPINFMAIL"
      Tab(4).Control(5)=   "chkUpdate_ROPINFMAIL"
      Tab(4).Control(6)=   "lblUpdate_ROPINFGNAT"
      Tab(4).Control(7)=   "libDossier_ROPDOSGUSR"
      Tab(4).Control(8)=   "libDossier_ROPDOSIUSR"
      Tab(4).Control(9)=   "libDossier_ROPDOSID"
      Tab(4).ControlCount=   10
      Begin VB.Frame fraDossier 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7900
         Left            =   -74800
         TabIndex        =   82
         Top             =   1110
         Visible         =   0   'False
         Width           =   10950
         Begin VB.ListBox lstUpdate_ROPDOSQUAL 
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   720
            Sorted          =   -1  'True
            TabIndex        =   157
            Top             =   4200
            Visible         =   0   'False
            Width           =   4600
         End
         Begin VB.Frame fraDossier_cmd 
            BackColor       =   &H00F0F0F0&
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
            Height          =   5292
            Left            =   9480
            TabIndex        =   127
            Top             =   120
            Width           =   1332
            Begin VB.CommandButton cmdDossier_Mail 
               BackColor       =   &H0080FFFF&
               Caption         =   "Envoyer Mail"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   648
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   132
               Top             =   120
               Width           =   1092
            End
            Begin VB.CommandButton cmdDossier_Print 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Imprimer"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   648
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   131
               Top             =   840
               Width           =   1092
            End
            Begin VB.CommandButton cmdUpdate_01 
               BackColor       =   &H0080C0FF&
               Caption         =   "Ajouter une note"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   612
               Left            =   120
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   130
               Top             =   2400
               Width           =   1095
            End
            Begin VB.CommandButton cmdUpdate_05 
               BackColor       =   &H0080C0FF&
               Caption         =   "Ajouter une pièce jointe"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   612
               Left            =   120
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   129
               Top             =   3120
               Width           =   1095
            End
            Begin VB.CommandButton cmdUpdate_02 
               BackColor       =   &H0080C0FF&
               Caption         =   "Ajouter une action"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   612
               Left            =   120
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   128
               Top             =   1680
               Width           =   1095
            End
         End
         Begin VB.CommandButton cmdDossier_Ok_Close 
            BackColor       =   &H00E0FFE0&
            Caption         =   "Enregistrer + fermer processus"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   732
            Left            =   9480
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   5520
            Width           =   1332
         End
         Begin VB.CommandButton cmdDossier_Ok 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Enregistrer"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   612
            Left            =   9480
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   6360
            Width           =   1332
         End
         Begin VB.CommandButton cmdDossier_Quit 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Abandonner"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   648
            Left            =   9480
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   7050
            Width           =   1332
         End
         Begin TabDlg.SSTab tabDossier 
            Height          =   7650
            Left            =   90
            TabIndex        =   86
            Top             =   180
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   13494
            _Version        =   393216
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   420
            BackColor       =   14737632
            ForeColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Dossier"
            TabPicture(0)   =   "DROPI.frx":0098
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fraDossier_B"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "cmdUpdate_Dossier"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Suivi"
            TabPicture(1)   =   "DROPI.frx":00B4
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "txtDetail"
            Tab(1).Control(1)=   "fgDetail"
            Tab(1).ControlCount=   2
            TabCaption(2)   =   "Tab 2"
            TabPicture(2)   =   "DROPI.frx":00D0
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "cmdUpdate_12"
            Tab(2).Control(1)=   "cmdUpdate_32"
            Tab(2).Control(2)=   "cmdUpdate_22"
            Tab(2).Control(3)=   "cmdUpdate"
            Tab(2).Control(4)=   "fraUpdate_B"
            Tab(2).ControlCount=   5
            TabCaption(3)   =   "Tab 3"
            TabPicture(3)   =   "DROPI.frx":00EC
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "lstUpdate_ROPINFMAIL_CC"
            Tab(3).Control(1)=   "lstUpdate_ROPINFMAIL_CC_Display"
            Tab(3).Control(2)=   "lstUpdate_ROPINFMAIL_Display"
            Tab(3).Control(3)=   "fraUpdate_PJ"
            Tab(3).Control(4)=   "lstUpdate_ROPINFMAIL"
            Tab(3).Control(5)=   "libUpdate_ROPINFMAIL_CC"
            Tab(3).Control(6)=   "libUpdate_ROPINFMAIL"
            Tab(3).ControlCount=   7
            Begin VB.ListBox lstUpdate_ROPINFMAIL_CC 
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2460
               Left            =   -74910
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   187
               Top             =   4890
               Width           =   4395
            End
            Begin VB.ListBox lstUpdate_ROPINFMAIL_CC_Display 
               BackColor       =   &H00FFFFF0&
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2370
               Left            =   -70275
               TabIndex        =   186
               Top             =   4950
               Width           =   4050
            End
            Begin VB.CommandButton cmdUpdate_12 
               BackColor       =   &H0080C0FF&
               Caption         =   "Modifier"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   700
               Left            =   -69600
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   177
               Top             =   1200
               Visible         =   0   'False
               Width           =   1100
            End
            Begin VB.CommandButton cmdUpdate_32 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Annuler"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   700
               Left            =   -67080
               Style           =   1  'Graphical
               TabIndex        =   176
               Top             =   1200
               Visible         =   0   'False
               Width           =   1100
            End
            Begin VB.CommandButton cmdUpdate_22 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Clôturer"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   700
               Left            =   -68330
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   175
               Top             =   1200
               Visible         =   0   'False
               Width           =   1100
            End
            Begin VB.ComboBox cmdUpdate 
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   330
               Left            =   -69600
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   145
               Top             =   720
               Width           =   3732
            End
            Begin VB.ListBox lstUpdate_ROPINFMAIL_Display 
               BackColor       =   &H00F0FFF0&
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3000
               Left            =   -70260
               TabIndex        =   174
               Top             =   915
               Width           =   4044
            End
            Begin VB.Frame fraUpdate_B 
               BackColor       =   &H00E0FFFF&
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   7092
               Left            =   -75000
               TabIndex        =   162
               Top             =   360
               Width           =   9200
               Begin VB.ComboBox txtUpdate_ROPINFSTA 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C000C0&
                  Height          =   312
                  Left            =   3600
                  Locked          =   -1  'True
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   180
                  Top             =   360
                  Width           =   1452
               End
               Begin VB.TextBox txtUpdate_ROPINFGUO 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   8040
                  TabIndex        =   166
                  Top             =   1680
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.ComboBox txtUpdate_ROPINFGNAT 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C000C0&
                  Height          =   312
                  Left            =   1560
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   165
                  Top             =   360
                  Width           =   1932
               End
               Begin VB.ComboBox txtUpdate_ROPINFGUSR 
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
                  Left            =   1560
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   164
                  Top             =   840
                  Width           =   3492
               End
               Begin VB.TextBox txtUpdate_ROPINFGTXT 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   3972
                  Left            =   240
                  MaxLength       =   1024
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   163
                  Top             =   2280
                  Width           =   8652
               End
               Begin MSComCtl2.DTPicker txtUpdate_ROPINFGECH 
                  Height          =   300
                  Left            =   1560
                  TabIndex        =   167
                  Top             =   1440
                  Width           =   1332
                  _ExtentX        =   2355
                  _ExtentY        =   529
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
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
                  Format          =   92340227
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin MSComCtl2.DTPicker txtUpdate_ROPINFGECH_Old 
                  Height          =   300
                  Left            =   5280
                  TabIndex        =   168
                  Top             =   1680
                  Visible         =   0   'False
                  Width           =   1332
                  _ExtentX        =   2355
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   16777215
                  CalendarForeColor=   0
                  CalendarTitleBackColor=   8421504
                  CalendarTitleForeColor=   16777215
                  CalendarTrailingForeColor=   12632256
                  CustomFormat    =   "dd  MM yyy"
                  Format          =   92340227
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin VB.Label lblUpdate_ROPINFGUO 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Durée (HH.MM)"
                  Height          =   252
                  Left            =   6720
                  TabIndex        =   172
                  Top             =   1680
                  Visible         =   0   'False
                  Width           =   1212
               End
               Begin VB.Label lblUpdate_ROPINFGECH 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Echéance"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   240
                  TabIndex        =   171
                  Top             =   1560
                  Width           =   1212
               End
               Begin VB.Label lblUpdate_ROPINFGUSR 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Responsable"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   240
                  TabIndex        =   170
                  Top             =   960
                  Width           =   1092
               End
               Begin VB.Label libDossier_ROPINFID 
                  BackColor       =   &H00E0F0FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "libUpdate_ROPINFID"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   240
                  TabIndex        =   169
                  Top             =   6600
                  Width           =   8652
               End
            End
            Begin VB.ComboBox cmdUpdate_Dossier 
               BackColor       =   &H00FF00FF&
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   360
               Left            =   6000
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   144
               Top             =   480
               Width           =   2892
            End
            Begin VB.TextBox txtDetail 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1260
               Left            =   -74400
               MultiLine       =   -1  'True
               TabIndex        =   125
               Text            =   "DROPI.frx":0108
               Top             =   1320
               Visible         =   0   'False
               Width           =   6732
            End
            Begin VB.Frame fraUpdate_PJ 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Pièce Jointe"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7095
               Left            =   -68685
               TabIndex        =   106
               Top             =   5175
               Visible         =   0   'False
               Width           =   8715
               Begin RichTextLib.RichTextBox rtfPJ 
                  Height          =   2712
                  Left            =   3240
                  TabIndex        =   178
                  TabStop         =   0   'False
                  Top             =   3480
                  Width           =   5028
                  _ExtentX        =   8864
                  _ExtentY        =   4789
                  _Version        =   393217
                  BackColor       =   14737632
                  HideSelection   =   0   'False
                  ScrollBars      =   3
                  AutoVerbMenu    =   -1  'True
                  TextRTF         =   $"DROPI.frx":0112
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
               Begin VB.DriveListBox DriveListBox 
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
                  Left            =   120
                  TabIndex        =   109
                  Top             =   360
                  Width           =   4000
               End
               Begin VB.DirListBox dirListBox 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1992
                  Left            =   120
                  TabIndex        =   108
                  Top             =   960
                  Width           =   4000
               End
               Begin VB.FileListBox filDoc 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00008000&
                  Height          =   2400
                  Left            =   4440
                  Pattern         =   "*.doc;*.pdf;*.rtf;*.xls;*.txt"
                  TabIndex        =   107
                  Top             =   360
                  Width           =   4000
               End
               Begin VB.Label librtfPJ 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Click droit pour copier/coller ==>"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   360
                  TabIndex        =   181
                  Top             =   4440
                  Width           =   2772
               End
            End
            Begin VB.ListBox lstUpdate_ROPINFMAIL 
               BackColor       =   &H00D0FFD0&
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3180
               Left            =   -74880
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   104
               Top             =   810
               Width           =   4395
            End
            Begin VB.Frame fraDossier_B 
               BackColor       =   &H00F0FFF0&
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
               Height          =   7332
               Left            =   0
               TabIndex        =   87
               Top             =   240
               Width           =   9060
               Begin VB.TextBox txtUpdate_ROPDOSGTXT 
                  BackColor       =   &H00FAFFFA&
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   2172
                  Left            =   120
                  Locked          =   -1  'True
                  MaxLength       =   1024
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   156
                  Top             =   2640
                  Width           =   8892
               End
               Begin VB.Frame fraDossier_C 
                  BackColor       =   &H00E0FFE0&
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2532
                  Left            =   120
                  TabIndex        =   133
                  Top             =   80
                  Width           =   8892
                  Begin VB.ComboBox txtUpdate_ROPDOSSTA 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   7.5
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000C0&
                     Height          =   312
                     Left            =   1560
                     Locked          =   -1  'True
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   179
                     Top             =   240
                     Width           =   2292
                  End
                  Begin VB.ComboBox txtUpdate_ROPDOSQUAL 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   7.5
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C00000&
                     Height          =   312
                     Left            =   5880
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   161
                     Top             =   2040
                     Width           =   2800
                  End
                  Begin VB.ComboBox txtUpdate_ROPDOSGNAT 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   7.5
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C00000&
                     Height          =   312
                     Left            =   1560
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   158
                     Top             =   2160
                     Width           =   1812
                  End
                  Begin VB.ComboBox txtUpdate_ROPDOSGUSR 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   7.5
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   312
                     Left            =   5880
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   139
                     Top             =   1080
                     Width           =   2800
                  End
                  Begin VB.ComboBox txtUpdate_ROPDOSIUSR 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   7.5
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   312
                     Left            =   1560
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   135
                     Top             =   1080
                     Width           =   2800
                  End
                  Begin MSComCtl2.DTPicker txtUpdate_ROPDOSIAMJ 
                     Height          =   300
                     Left            =   1560
                     TabIndex        =   137
                     Top             =   1680
                     Width           =   1332
                     _ExtentX        =   2355
                     _ExtentY        =   529
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "@Arial Unicode MS"
                        Size            =   7.5
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     CalendarBackColor=   16777215
                     CalendarForeColor=   12582912
                     CalendarTitleBackColor=   8421504
                     CalendarTitleForeColor=   16777215
                     CalendarTrailingForeColor=   12632256
                     CustomFormat    =   "dd  MM yyy"
                     Format          =   92340227
                     CurrentDate     =   38699.44875
                     MaxDate         =   401768
                     MinDate         =   36526.4425347222
                  End
                  Begin MSComCtl2.DTPicker txtUpdate_ROPDOSGECH 
                     Height          =   300
                     Left            =   5880
                     TabIndex        =   140
                     Top             =   1560
                     Width           =   1212
                     _ExtentX        =   2143
                     _ExtentY        =   529
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "@Arial Unicode MS"
                        Size            =   7.5
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     CalendarBackColor=   16777215
                     CalendarForeColor=   12582912
                     CalendarTitleBackColor=   8421504
                     CalendarTitleForeColor=   16777215
                     CalendarTrailingForeColor=   12632256
                     CustomFormat    =   "dd  MM yyy"
                     Format          =   92340227
                     CurrentDate     =   38699.44875
                     MaxDate         =   401768
                     MinDate         =   36526.4425347222
                  End
                  Begin VB.Label lblUpdate_ROPDOSQUAL 
                     BackColor       =   &H00E0FFE0&
                     Caption         =   "Qualification"
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   7.5
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Left            =   4680
                     TabIndex        =   160
                     Top             =   2040
                     Width           =   1092
                  End
                  Begin VB.Label lblUpdate_ROPDOSGNAT 
                     BackColor       =   &H00E0FFE0&
                     Caption         =   "Nature"
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   7.5
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Left            =   120
                     TabIndex        =   159
                     Top             =   2160
                     Width           =   732
                  End
                  Begin VB.Label libUpdate_ROPDOSGSRV 
                     BackColor       =   &H00E0FFE0&
                     Caption         =   "libUpdate_ROPDOSGSRV"
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   7.5
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Left            =   5880
                     TabIndex        =   154
                     Top             =   720
                     Width           =   2892
                  End
                  Begin VB.Label libUpdate_ROPDOSISRV 
                     BackColor       =   &H00E0FFE0&
                     Caption         =   "libUpdate_ROPDOSISRV"
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   7.5
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Left            =   1560
                     TabIndex        =   153
                     Top             =   720
                     Width           =   3132
                  End
                  Begin VB.Label lblUpdate_ROPDOSGECH 
                     BackColor       =   &H00E0FFE0&
                     Caption         =   "Echéance"
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   7.5
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Left            =   4680
                     TabIndex        =   141
                     Top             =   1560
                     Width           =   852
                  End
                  Begin VB.Label lblUpdate_ROPDOSGUSR 
                     BackColor       =   &H00E0FFE0&
                     Caption         =   "Gestionnaire"
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   7.5
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Left            =   4680
                     TabIndex        =   138
                     Top             =   1200
                     Width           =   1092
                  End
                  Begin VB.Label lblUpdate_ROPDOSIAMJ 
                     BackColor       =   &H00E0FFE0&
                     Caption         =   "date du constat"
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   7.5
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Left            =   120
                     TabIndex        =   136
                     Top             =   1680
                     Width           =   1332
                  End
                  Begin VB.Label lblUpdate_ROPDOSIUSR 
                     BackColor       =   &H00E0FFE0&
                     Caption         =   "Initiateur"
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   7.5
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Left            =   240
                     TabIndex        =   134
                     Top             =   1080
                     Width           =   852
                  End
               End
               Begin VB.ComboBox txtUpdate_ROPDOSGGRA 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   1320
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   95
                  Top             =   5500
                  Width           =   2800
               End
               Begin VB.TextBox txtUpdate_ROPDOSGCOU 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   324
                  Left            =   1320
                  TabIndex        =   94
                  Top             =   6000
                  Width           =   1692
               End
               Begin VB.ComboBox txtUpdate_ROPDOSGPRI 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   1320
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   93
                  Top             =   5000
                  Width           =   2800
               End
               Begin VB.ComboBox txtUpdate_ROPDOSGPRV 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   1320
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   92
                  Top             =   6500
                  Width           =   1572
               End
               Begin VB.TextBox txtUpdate_ROPDOSXID 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   324
                  Left            =   5880
                  TabIndex        =   91
                  Top             =   6000
                  Width           =   2800
               End
               Begin VB.TextBox txtUpdate_ROPDOSIREF 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   324
                  Left            =   5880
                  TabIndex        =   90
                  Top             =   6500
                  Width           =   2800
               End
               Begin VB.ComboBox txtUpdate_ROPDOSXAPP 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   5880
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   89
                  Top             =   5500
                  Width           =   2800
               End
               Begin VB.ComboBox txtUpdate_ROPDOSXDOM 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   5880
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   88
                  Top             =   5000
                  Width           =   2800
               End
               Begin VB.Label lblUpdate_ROPDOSGGRA 
                  BackColor       =   &H00F0FFF0&
                  Caption         =   "Gravité"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   240
                  TabIndex        =   103
                  Top             =   5600
                  Width           =   852
               End
               Begin VB.Label lblUpdate_ROPDOSGCOU 
                  BackColor       =   &H00F0FFF0&
                  Caption         =   "Coût €"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   240
                  TabIndex        =   102
                  Top             =   6100
                  Width           =   852
               End
               Begin VB.Label lblUpdate_ROPDOSGPRI 
                  BackColor       =   &H00F0FFF0&
                  Caption         =   "Priorité"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   240
                  TabIndex        =   101
                  Top             =   5100
                  Width           =   972
               End
               Begin VB.Label lblUpdate_ROPDOSIREF 
                  BackColor       =   &H00F0FFF0&
                  Caption         =   "Réf interne"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   4596
                  TabIndex        =   100
                  Top             =   6600
                  Width           =   972
               End
               Begin VB.Label lblUpdate_ROPDOSXID 
                  BackColor       =   &H00F0FFF0&
                  Caption         =   "Réf externe"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   4596
                  TabIndex        =   99
                  Top             =   6100
                  Width           =   1092
               End
               Begin VB.Label lblUpdate_ROPDOSXAPP 
                  BackColor       =   &H00F0FFF0&
                  Caption         =   "Application"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   4560
                  TabIndex        =   98
                  Top             =   5600
                  Width           =   972
               End
               Begin VB.Label lblUpdate_ROPDOSXDOM 
                  BackColor       =   &H00F0FFF0&
                  Caption         =   "Domaine"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   4560
                  TabIndex        =   97
                  Top             =   5100
                  Width           =   732
               End
               Begin VB.Label libDossier_ROPDOSUUSR 
                  BackColor       =   &H00F0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "saisie"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Left            =   120
                  TabIndex        =   96
                  Top             =   6960
                  Width           =   8532
               End
            End
            Begin MSFlexGridLib.MSFlexGrid fgDetail 
               Height          =   7080
               Left            =   -74940
               TabIndex        =   126
               Top             =   480
               Width           =   9120
               _ExtentX        =   16087
               _ExtentY        =   12488
               _Version        =   393216
               Rows            =   1
               RowHeightMin    =   250
               BackColor       =   16777215
               ForeColor       =   12582912
               BackColorFixed  =   13693183
               ForeColorFixed  =   8388608
               BackColorSel    =   14737632
               ForeColorSel    =   8388736
               BackColorBkg    =   16777210
               WordWrap        =   -1  'True
               AllowBigSelection=   0   'False
               FocusRect       =   2
               HighLight       =   0
               GridLinesFixed  =   1
               AllowUserResizing=   3
               FormatString    =   "<                                                                        |"
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
            Begin VB.Label libUpdate_ROPINFMAIL_CC 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "COPIE à :"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   -71175
               TabIndex        =   185
               Top             =   4365
               Visible         =   0   'False
               Width           =   1710
            End
            Begin VB.Label libUpdate_ROPINFMAIL 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "ENVOYER à :"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -71385
               TabIndex        =   184
               Top             =   330
               Visible         =   0   'False
               Width           =   1950
            End
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H009BCFFF&
         Caption         =   "Frame1"
         Height          =   2172
         Left            =   -73440
         TabIndex        =   146
         Top             =   2760
         Width           =   8892
         Begin VB.ListBox lstUpdate_Modèle 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   151
            Top             =   1560
            Visible         =   0   'False
            Width           =   3012
         End
         Begin VB.CheckBox chkUpdate_ROPINFGPRV 
            BackColor       =   &H00E0FFFF&
            Caption         =   "modifiable uniquement par l'auteur"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   150
            Top             =   960
            Width           =   3132
         End
         Begin VB.TextBox txtUpdate_ROPINFIDTL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   240
            TabIndex        =   148
            Text            =   "123"
            Top             =   840
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblUpdate_ROPINFIDTL 
            BackColor       =   &H00E0FFFF&
            Caption         =   "attendre la fin de l'action"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   960
            TabIndex        =   149
            Top             =   840
            Visible         =   0   'False
            Width           =   2412
         End
         Begin VB.Label txtUpdate_ROPINFGTXT_0 
            BackColor       =   &H00D0FFD0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "txtUpdate_ROPINFGTXT_0"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   0
            TabIndex        =   147
            Top             =   360
            Width           =   8772
            WordWrap        =   -1  'True
         End
      End
      Begin VB.ListBox X_cmdUpdate_Dossier 
         BackColor       =   &H00F0FFF0&
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         ItemData        =   "DROPI.frx":0189
         Left            =   -73560
         List            =   "DROPI.frx":0190
         Sorted          =   -1  'True
         TabIndex        =   143
         Top             =   480
         Width           =   3132
      End
      Begin VB.ListBox X_cmdUpdate 
         BackColor       =   &H00F0FFF0&
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   -69960
         TabIndex        =   142
         Top             =   600
         Width           =   3564
      End
      Begin VB.ListBox libDossier_ROPINFGTXT 
         BackColor       =   &H00F0FFF0&
         Height          =   255
         Left            =   -63840
         TabIndex        =   124
         Top             =   1680
         Width           =   2724
      End
      Begin VB.Frame fraUpdate_ROPINFMAIL 
         BackColor       =   &H00E0FFFF&
         Caption         =   "Destinataires du mail de suivi"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -64440
         TabIndex        =   113
         Top             =   3720
         Width           =   3855
         Begin VB.CheckBox chkUpdate_ROPINFMAIL_U 
            BackColor       =   &H00E0FFFF&
            Caption         =   "moi"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   118
            Top             =   1440
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate_ROPINFMAIL_A 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Resp action"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   117
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate_ROPINFMAIL_I 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Initiateur"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   116
            Top             =   600
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate_ROPINFMAIL_P 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Resp processus"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   115
            Top             =   900
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate_ROPINFMAIL_D 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Gest du dossier"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   1440
            TabIndex        =   114
            Top             =   300
            Width           =   2295
         End
         Begin VB.Label lblUpdate_ROPINFMAIL_U 
            BackColor       =   &H00E0FFFF&
            Caption         =   "au suivant....."
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   1500
            Width           =   1095
         End
         Begin VB.Label lblUpdate_ROPINFMAIL_A 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Resp Action"
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblUpdate_ROPINFMAIL_P 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Gest Processus"
            Height          =   255
            Left            =   120
            TabIndex        =   121
            Top             =   900
            Width           =   1215
         End
         Begin VB.Label lblUpdate_ROPINFMAIL_I 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Initiateur"
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblUpdate_ROPINFMAIL_D 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Superviseur"
            Height          =   255
            Left            =   120
            TabIndex        =   119
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkUpdate_ROPINFMAIL 
         BackColor       =   &H00F0FFFF&
         Caption         =   "Suivi mail automatique"
         Height          =   255
         Left            =   -63240
         TabIndex        =   105
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Frame fraExport 
         BackColor       =   &H00D0FFD0&
         Caption         =   "Export => C:\temp\DROPI.xlsx et C:\temp\DROPI_B2.xlsx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   -66105
         TabIndex        =   67
         Top             =   1170
         Visible         =   0   'False
         Width           =   5655
         Begin VB.CheckBox chkExport_ROPINFGTXT 
            BackColor       =   &H00D0FFD0&
            Caption         =   "exporter la description"
            Height          =   255
            Left            =   1560
            TabIndex        =   75
            Top             =   1560
            Width           =   3375
         End
         Begin VB.CheckBox chkExport_ROPDOSGNAT 
            BackColor       =   &H00D0FFD0&
            Caption         =   "exporter uniquement les indicents"
            Height          =   255
            Left            =   1560
            TabIndex        =   70
            Top             =   1080
            Value           =   1  'Checked
            Width           =   3375
         End
         Begin VB.CommandButton cmdExport_Quit 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Abandonner"
            Height          =   525
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CommandButton cmdExport_Ok 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Exporter"
            Height          =   525
            Left            =   3720
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   2280
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker txtExport_AMJMIN 
            Height          =   300
            Left            =   1560
            TabIndex        =   71
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
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
            Format          =   92340227
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin MSComCtl2.DTPicker txtExport_AMJMAX 
            Height          =   300
            Left            =   3720
            TabIndex        =   72
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
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
            Format          =   92340227
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin VB.Label lblExport_AMJMAX 
            BackColor       =   &H00D0FFD0&
            Caption         =   "au"
            Height          =   255
            Left            =   3120
            TabIndex        =   74
            Top             =   600
            Width           =   495
         End
         Begin VB.Label lblExport_AMJMIN 
            BackColor       =   &H00D0FFD0&
            Caption         =   "Dossiers crées du"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.ListBox lstW 
         BackColor       =   &H80000001&
         Height          =   3375
         Left            =   -63720
         Sorted          =   -1  'True
         TabIndex        =   50
         Top             =   4680
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Frame fraAut 
         BackColor       =   &H00C0F0FF&
         Height          =   9255
         Left            =   -74880
         TabIndex        =   29
         Top             =   480
         Width           =   14535
         Begin VB.ListBox lstAut_Usr 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6984
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   5025
         End
         Begin VB.Frame fraAut_Update 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Mise à jour"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8655
            Left            =   5400
            TabIndex        =   30
            Top             =   360
            Width           =   8895
            Begin VB.CheckBox chkAut_ROPDOSGUSR_I 
               BackColor       =   &H00E0FFFF&
               Caption         =   "réception mail Incidents significatifs"
               Height          =   492
               Left            =   3120
               TabIndex        =   183
               Top             =   840
               Visible         =   0   'False
               Width           =   2652
            End
            Begin VB.CheckBox chkAut_ROPDOSGUSR_E 
               BackColor       =   &H00E0FFFF&
               Caption         =   "exportation des dossiers =>.xlsx"
               Height          =   255
               Left            =   3120
               TabIndex        =   182
               Top             =   480
               Width           =   3012
            End
            Begin VB.TextBox txtAut_ROPDOSGUSR_SRV 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   240
               MaxLength       =   3
               TabIndex        =   79
               Text            =   "S99"
               Top             =   2280
               Width           =   612
            End
            Begin VB.CheckBox chkAut_ROPDOSGUSR_Q 
               BackColor       =   &H00E0FFFF&
               Caption         =   "gestion des qualifications"
               Height          =   255
               Left            =   240
               TabIndex        =   51
               Top             =   1680
               Width           =   2535
            End
            Begin VB.ListBox lstAut_ROPDOSGUSR 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   690
               Left            =   240
               TabIndex        =   47
               Top             =   6600
               Width           =   3105
            End
            Begin VB.Frame fraAut_ROPDOSISRV 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Habilitation utilisateur / service"
               Height          =   3015
               Left            =   240
               TabIndex        =   40
               Top             =   3360
               Width           =   3135
               Begin VB.OptionButton optAut_ROPDOSISRV_Z 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "sans lien avec le service"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   48
                  Top             =   2520
                  Width           =   2415
               End
               Begin VB.OptionButton optAut_ROPDOSISRV_I 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Inspection"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   46
                  Top             =   2160
                  Width           =   2415
               End
               Begin VB.OptionButton optAut_ROPDOSISRV_X 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "X sans habilitation"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   45
                  Top             =   1800
                  Width           =   2415
               End
               Begin VB.OptionButton optAut_ROPDOSISRV_C 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Collaboteur"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   44
                  Top             =   1440
                  Width           =   2415
               End
               Begin VB.OptionButton optAut_ROPDOSISRV_D 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Délégation de gestion"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   43
                  Top             =   1080
                  Width           =   2415
               End
               Begin VB.OptionButton optAut_ROPDOSISRV_R 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Responsable"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   42
                  Top             =   720
                  Width           =   2415
               End
               Begin VB.OptionButton optAut_ROPDOSISRV_H 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Hierarchie"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   41
                  Top             =   360
                  Width           =   2415
               End
            End
            Begin VB.ListBox lstAut_ROPDOSISRV 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4050
               Left            =   3720
               TabIndex        =   39
               Top             =   3360
               Width           =   4905
            End
            Begin VB.CheckBox chkAut_ROPDOSGUSR_H 
               BackColor       =   &H00E0FFFF&
               Caption         =   "gestion des habilitations"
               Height          =   255
               Left            =   240
               TabIndex        =   38
               Top             =   1380
               Width           =   2535
            End
            Begin VB.CheckBox chkAut_ROPDOSGUSR_P 
               BackColor       =   &H00E0FFFF&
               Caption         =   "gestion du paramétrage"
               Height          =   255
               Left            =   240
               TabIndex        =   37
               Top             =   1080
               Width           =   2535
            End
            Begin VB.CheckBox chkAut_ROPINFGUSR 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Responsable d'action"
               Height          =   255
               Left            =   240
               TabIndex        =   35
               Top             =   780
               Width           =   2535
            End
            Begin VB.CheckBox chkAut_ROPDOSGUSR 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Gestionnaire de dossier"
               Height          =   255
               Left            =   240
               TabIndex        =   34
               Top             =   480
               Width           =   2535
            End
            Begin VB.CommandButton cmdAut_Update_Ok 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Enregistrer"
               Height          =   765
               Left            =   7320
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   480
               Width           =   1452
            End
            Begin VB.CommandButton cmdAut_Update_Quit 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Abandonner"
               Height          =   765
               Left            =   7320
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   1440
               Width           =   1452
            End
            Begin VB.Label libAut_ROPDOSGUSR_SRV 
               BackColor       =   &H00E0FFFF&
               Caption         =   "préciser le service"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1080
               TabIndex        =   80
               Top             =   2280
               Width           =   3492
            End
            Begin VB.Label lblAut_ROPDOSGUSR_Mail 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblAut_ROPDOSGUSR_Mail"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   240
               TabIndex        =   49
               Top             =   2760
               Width           =   4332
            End
         End
      End
      Begin VB.Frame fraParam 
         Height          =   9375
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   14535
         Begin TabDlg.SSTab SSTab2 
            Height          =   7215
            Left            =   120
            TabIndex        =   56
            Top             =   120
            Width           =   14295
            _ExtentX        =   25215
            _ExtentY        =   12726
            _Version        =   393216
            TabHeight       =   520
            ForeColor       =   8388736
            TabCaption(0)   =   "DOMAINES / APPLICATIONS"
            TabPicture(0)   =   "DROPI.frx":01A9
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lblParam_ROPDOSXDOM"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "lblParam_ROPDOSXAPP"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "lblParam_ROPINFGTXT"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "libROPDOSMAIL"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "lstParam_ROPDOSXDOM"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "lstParam_ROPDOSXAPP"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "lstParam_ROPINFGTXT"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "txtROPDOSMAIL"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).ControlCount=   8
            TabCaption(1)   =   "QUALIFICATION BALE II"
            TabPicture(1)   =   "DROPI.frx":01C5
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lstParam_ROPDOSQUAL"
            Tab(1).Control(1)=   "lstParam_ROPDOSQUALB2"
            Tab(1).Control(2)=   "lblParam_ROPDOSQUAL"
            Tab(1).Control(3)=   "lblParam_ROPDOSQUALB2"
            Tab(1).ControlCount=   4
            TabCaption(2)   =   "paramétrage des services"
            TabPicture(2)   =   "DROPI.frx":01E1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "lstParam_ROPDOSGUSR"
            Tab(2).Control(1)=   "lstParam_ROPDOSISRV"
            Tab(2).Control(2)=   "lblParam_ROPDOSISRV"
            Tab(2).ControlCount=   3
            Begin VB.TextBox txtROPDOSMAIL 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   900
               Left            =   4560
               MultiLine       =   -1  'True
               TabIndex        =   188
               Text            =   "DROPI.frx":01FD
               Top             =   5970
               Visible         =   0   'False
               Width           =   9120
            End
            Begin VB.ListBox lstParam_ROPDOSGUSR 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4380
               Left            =   -66720
               TabIndex        =   78
               Top             =   1080
               Width           =   5265
            End
            Begin VB.ListBox lstParam_ROPDOSISRV 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4380
               Left            =   -74640
               TabIndex        =   76
               Top             =   1080
               Width           =   7065
            End
            Begin VB.ListBox lstParam_ROPDOSQUAL 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4380
               Left            =   -68520
               TabIndex        =   65
               Top             =   1080
               Width           =   7425
            End
            Begin VB.ListBox lstParam_ROPDOSQUALB2 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4380
               Left            =   -74640
               TabIndex        =   63
               Top             =   1080
               Width           =   5265
            End
            Begin VB.ListBox lstParam_ROPINFGTXT 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4380
               Left            =   9000
               TabIndex        =   61
               Top             =   1080
               Width           =   4665
            End
            Begin VB.ListBox lstParam_ROPDOSXAPP 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4380
               Left            =   4440
               TabIndex        =   59
               Top             =   1080
               Width           =   3700
            End
            Begin VB.ListBox lstParam_ROPDOSXDOM 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4380
               Left            =   120
               TabIndex        =   57
               Top             =   1080
               Width           =   3700
            End
            Begin VB.Label libROPDOSMAIL 
               BackColor       =   &H00FF80FF&
               Caption         =   $"DROPI.frx":020B
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1050
               Left            =   165
               TabIndex        =   189
               Top             =   5895
               Visible         =   0   'False
               Width           =   4140
            End
            Begin VB.Label lblParam_ROPDOSISRV 
               Caption         =   "SERVICES"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   -72240
               TabIndex        =   77
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label lblParam_ROPDOSQUAL 
               Caption         =   "Référentiel R O"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   -66120
               TabIndex        =   66
               Top             =   600
               Width           =   2535
            End
            Begin VB.Label lblParam_ROPDOSQUALB2 
               Caption         =   "Référentiel Bâle II"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   -73200
               TabIndex        =   64
               Top             =   600
               Width           =   2535
            End
            Begin VB.Label lblParam_ROPINFGTXT 
               Caption         =   "Libellé action"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   10080
               TabIndex        =   62
               Top             =   600
               Width           =   2055
            End
            Begin VB.Label lblParam_ROPDOSXAPP 
               Caption         =   "Applications"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   5280
               TabIndex        =   60
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label lblParam_ROPDOSXDOM 
               Caption         =   "Domaines"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   960
               TabIndex        =   58
               Top             =   600
               Width           =   1335
            End
         End
         Begin VB.Frame fraParam_Update 
            BackColor       =   &H00D0FFD0&
            Caption         =   "Mise à jour"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   120
            TabIndex        =   23
            Top             =   7320
            Width           =   14295
            Begin VB.CommandButton cmdParam_Update_Quit 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Abandonner"
               Height          =   765
               Left            =   10320
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   960
               Width           =   1335
            End
            Begin VB.CommandButton cmdParam_Update_Ok 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Enregistrer"
               Height          =   765
               Left            =   12600
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox txtParam_BIATABTXT 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   7920
               TabIndex        =   26
               Top             =   480
               Width           =   6015
            End
            Begin VB.TextBox txtParam_BIATABK2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5640
               TabIndex        =   25
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtParam_BIATABK1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   24
               Top             =   480
               Width           =   1935
            End
         End
      End
      Begin VB.Frame fraTab0 
         BackColor       =   &H00F0FFFF&
         Height          =   9525
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   14640
         Begin VB.CommandButton cmdSelect_New 
            BackColor       =   &H0080C0FF&
            Caption         =   "Saisir une fiche R.O."
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   612
            Left            =   11520
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   155
            Top             =   720
            Width           =   1095
         End
         Begin VB.Frame fraSelect 
            BackColor       =   &H00F0FFFF&
            Height          =   8052
            Left            =   100
            TabIndex        =   20
            Top             =   1440
            Width           =   14415
            Begin VB.TextBox txtSelect_txt 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1260
               Left            =   5040
               MultiLine       =   -1  'True
               TabIndex        =   152
               Text            =   "DROPI.frx":02E4
               Top             =   2520
               Visible         =   0   'False
               Width           =   6732
            End
            Begin MSFlexGridLib.MSFlexGrid fgSelect 
               Height          =   7680
               Left            =   120
               TabIndex        =   81
               Top             =   240
               Visible         =   0   'False
               Width           =   14160
               _ExtentX        =   24977
               _ExtentY        =   13547
               _Version        =   393216
               Rows            =   1
               Cols            =   5
               FixedCols       =   0
               RowHeightMin    =   750
               BackColor       =   15794175
               ForeColor       =   8388608
               BackColorFixed  =   8421376
               ForeColorFixed  =   16777215
               BackColorSel    =   14737632
               BackColorBkg    =   16777210
               WordWrap        =   -1  'True
               AllowBigSelection=   0   'False
               FocusRect       =   2
               HighLight       =   0
               GridLinesFixed  =   1
               AllowUserResizing=   3
               FormatString    =   $"DROPI.frx":02EE
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
            Left            =   11520
            Sorted          =   -1  'True
            TabIndex        =   12
            Text            =   "cboSelect_SQL"
            Top             =   240
            Width           =   2892
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00D0FFD0&
            Caption         =   "Rechercher"
            Height          =   645
            Left            =   12840
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   720
            Width           =   1572
         End
         Begin VB.Frame fraSelect_Options_1 
            BackColor       =   &H00E0D0D0&
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1440
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   11232
            Begin VB.CheckBox chkSelect_ROPDOSUAMJ 
               BackColor       =   &H00E0D0D0&
               Caption         =   "Créé>="
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   250
               Width           =   855
            End
            Begin VB.ComboBox txtSelect_ROPDOSGNAT 
               Height          =   288
               Left            =   6600
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   53
               Top             =   600
               Width           =   1575
            End
            Begin VB.ComboBox txtSelect_ROPDOSGPRV 
               Height          =   288
               Left            =   6600
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   1000
               Width           =   1575
            End
            Begin VB.CheckBox chkSelect_ROPDOSGUSR 
               BackColor       =   &H00E0D0D0&
               Caption         =   "Resp"
               Height          =   255
               Left            =   3000
               TabIndex        =   10
               Top             =   1050
               Width           =   735
            End
            Begin VB.TextBox txtSelect_ROPDOSGUSR 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   336
               Left            =   3840
               TabIndex        =   11
               Top             =   1000
               Width           =   2175
            End
            Begin VB.CheckBox chkSelect_ROPDOSGECH 
               BackColor       =   &H00E0D0D0&
               Caption         =   "éch <="
               Height          =   255
               Left            =   120
               TabIndex        =   5
               Top             =   1050
               Width           =   855
            End
            Begin VB.ComboBox txtSelect_ROPDOSSTA 
               Height          =   288
               Left            =   6600
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   200
               Width           =   1575
            End
            Begin VB.TextBox txtSelect_ROPINFGTXT 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   336
               Left            =   9720
               TabIndex        =   3
               Top             =   1000
               Width           =   1095
            End
            Begin VB.TextBox txtSelect_ROPDOSXID 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9720
               TabIndex        =   2
               Top             =   600
               Width           =   1095
            End
            Begin VB.ComboBox txtSelect_ROPDOSXAPP 
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
               Left            =   3000
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   600
               Width           =   3135
            End
            Begin VB.ComboBox txtSelect_ROPDOSXDOM 
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
               Left            =   3000
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   200
               Width           =   3135
            End
            Begin VB.TextBox txtSelect_ROPDOSID 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   336
               Left            =   9720
               TabIndex        =   1
               Top             =   200
               Width           =   1095
            End
            Begin MSComCtl2.DTPicker txtSelect_ROPDOSGECH_Max 
               Height          =   300
               Left            =   1080
               TabIndex        =   4
               Top             =   1000
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
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
               Format          =   92340227
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_ROPDOSUAMJ 
               Height          =   300
               Left            =   1080
               TabIndex        =   54
               Top             =   200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
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
               Format          =   92340227
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_ROPINFGTXT 
               BackColor       =   &H00E0D0D0&
               Caption         =   "texte à rechercher"
               Height          =   252
               Left            =   8280
               TabIndex        =   52
               Top             =   1080
               Width           =   1332
            End
            Begin VB.Label lblSelect_ROPDOSXID 
               BackColor       =   &H00E0D0D0&
               Caption         =   "N° CRI"
               Height          =   252
               Left            =   8760
               TabIndex        =   36
               Top             =   648
               Width           =   732
            End
            Begin VB.Label lblSelect_ROPDOSID 
               BackColor       =   &H00E0D0D0&
               Caption         =   "N° dossier"
               Height          =   252
               Left            =   8760
               TabIndex        =   21
               Top             =   252
               Width           =   852
            End
         End
      End
      Begin VB.Label lblUpdate_ROPINFGNAT 
         BackColor       =   &H00E0FFFF&
         Caption         =   "Nature"
         Height          =   252
         Left            =   -69360
         TabIndex        =   173
         Top             =   6000
         Width           =   612
      End
      Begin VB.Label libDossier_ROPDOSGUSR 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "libUpdate_ROPDOSGUSR"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   -63840
         TabIndex        =   112
         Top             =   3240
         Width           =   2772
      End
      Begin VB.Label libDossier_ROPDOSIUSR 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "libUpdate_ROPDOSIUSR"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   300
         Left            =   -63840
         TabIndex        =   111
         Top             =   2640
         Width           =   3012
      End
      Begin VB.Label libDossier_ROPDOSID 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   -63480
         TabIndex        =   110
         Top             =   2040
         Width           =   2172
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   14160
      Picture         =   "DROPI.frx":0405
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   732
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0FFFF&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label libRéférenceInterne 
      BackColor       =   &H00C0FFC0&
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
      Height          =   495
      Left            =   1200
      TabIndex        =   14
      Top             =   0
      Width           =   5535
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuExport_Service 
         Caption         =   "Export par SERVICE des dossiers en vie"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export =>C:\temp\DROPI.xlsx"
      End
      Begin VB.Menu mnuExport_Param 
         Caption         =   "Export =>C:\temp\DROPI_Param.xlsx"
      End
      Begin VB.Menu mnuExport_Migration 
         Caption         =   "Export_Migration"
      End
      Begin VB.Menu mnuContext_x1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuparam 
      Caption         =   "mnuParam"
      Visible         =   0   'False
      Begin VB.Menu mnuParam_DOMAINES 
         Caption         =   "                DOMAINES"
      End
      Begin VB.Menu mnuparam_APPLICATIONS 
         Caption         =   "               APLLICATIONS"
      End
      Begin VB.Menu mnuparam_LIBELLES 
         Caption         =   "                LIBELLES"
      End
      Begin VB.Menu mnuparam_ROPDOSQUAL 
         Caption         =   "            Référentiel RO"
      End
      Begin VB.Menu mnuparam_ROPDOSQUALB2 
         Caption         =   "          Référentiel Bâle II"
      End
      Begin VB.Menu mnuParam_ROPDOSISRV 
         Caption         =   "               SERVICES"
      End
      Begin VB.Menu mnuParam_Insert 
         Caption         =   "Ajouter un enregistrement"
      End
      Begin VB.Menu mnuParam_Update 
         Caption         =   " Modifier cet enregistrement"
      End
      Begin VB.Menu mnuParam_Delete 
         Caption         =   "Supprimer cet  enregistrement"
      End
      Begin VB.Menu mnuParam_Copy 
         Caption         =   "Copier cet  enregistrement"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint0_Dossier_All_XDOM 
         Caption         =   "Imprimer tous les dossiers (classés par domaine)"
      End
      Begin VB.Menu mnuPrint0_Dossier_All 
         Caption         =   "Imprimer tous les dossiers (classés par identifiant)"
      End
   End
End
Attribute VB_Name = "frmDROPI"
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
Dim intReturn As Integer, vDsys As Variant
Dim DROPI_Aut As typeAuthorization, mAPP_Menu As String
Dim blnTransaction As Boolean
Dim blnAuto As Boolean, blnAuto_Ok As Boolean
Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long
Dim wAmjMin7 As Long, wAmjMax7 As Long


Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnSetfocus As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean
Dim cmdSelect_Ok_Caption As String
Dim cmdSelect_SQL_K As String, cmdSelect_SQL_X1 As String
Dim xYROPDOS0 As typeYROPDOS0, meYROPDOS0 As typeYROPDOS0
Dim newYROPDOS0 As typeYROPDOS0, oldYROPDOS0 As typeYROPDOS0
Dim arrYROPDOS0() As typeYROPDOS0, arrYROPDOS0_Nb As Long, arrYROPDOS0_Max As Long, arrYROPDOS0_Index As Long
Dim selYROPDOS0() As typeYROPDOS0, selYROPDOS0_Nb As Long, selYROPDOS0_Max As Long, selYROPDOS0_Index As Long
Dim mailYROPDOS0 As typeYROPDOS0
Dim currentROPDOSID As String

'______________________________________________________________________

Dim fgDetail_FormatString As String, fgDetail_K As Integer
Dim fgDetail_RowDisplay As Integer, fgDetail_RowClick As Integer, fgDetail_ColClick As Integer
Dim fgDetail_ColorClick As Long, fgDetail_ColorDisplay As Long
Dim fgDetail_Sort1 As Integer, fgDetail_Sort2 As Integer
Dim fgDetail_SortAD As Integer, fgDetail_Sort1_Old As Integer
Dim fgDetail_arrIndex As Integer
Dim blnfgDetail_DisplayLine As Boolean
Dim cmdDétail_Ok_Caption As String
Dim cmdDétail_SQL_K As String
Dim xYROPINF0 As typeYROPINF0, meYROPINF0 As typeYROPINF0
Dim newYROPINF0 As typeYROPINF0, oldYROPINF0 As typeYROPINF0
Dim arrYROPINF0() As typeYROPINF0, arrYROPINF0_Nb As Long, arrYROPINF0_Max As Long, arrYROPINF0_Index As Long
Dim selYROPINF0() As typeYROPINF0
Dim zYROPINF0 As typeYROPINF0
Dim mailYROPINF0 As typeYROPINF0, mailYROPINF0_Suivant As typeYROPINF0
Dim savYROPINF0 As typeYROPINF0
Dim mUpdate_Nature As String, mUpdate_Action As String
'______________________________________________________________________

Dim cmdUpdate_K As String, cmdUpdate_Init_K As String, cmdUpdate_Fct As String
Dim mROPINFSTA_Value As String, mROPINFSTA_Set As String, mROPINFSTA_Where As String
Dim mROPINFSTAK_Set As String, mROPINFSTAD_Set As String, mROPINFSTAD_Where As String

Dim Dossier_Aut As typeAuthorization, Processus_Aut As typeAuthorization, Action_Aut As typeAuthorization, Memo_Aut As typeAuthorization
Dim False_Aut As typeAuthorization
Dim Processus_Index As Long, Action_Index As Long, Action_Suivante_Index As Long

Dim blnProcessus_EnCours As Boolean, blnProcessus_EnAlerte As Boolean
Dim blnAction_EnCours As Boolean, blnAction_EnAlerte As Boolean, blnAction_Valide As Boolean
Dim blnAction22et02 As Boolean

Dim blnParam_ROPDOSXDOM As Boolean, cmdParam_SQL_K As String
Dim oldParam As typeYBIATAB0, newParam As typeYBIATAB0
Dim blnSendMail As Boolean, mailSubject As String
Dim oldFileName As String, newFileName As String, newDirPath As String, newFileExtension As String

Dim cmdAut_SQL_K As String
Dim oldAut As typeYBIATAB0, newAut As typeYBIATAB0

Dim arrSelect_Update(50) As String, arrSelect_Update_Nb As Integer
Dim blnSelect_Update_EnCours As Boolean
Dim blnDossierModèle As Boolean, DossierModèle_ROPDOSID As Long
Dim blnDossierReprise As Boolean
Dim blnParam_ROPINFGTXT As Boolean, arrROPINFGTXT_BIATABK2() As String
Dim arrROPDOSISRV_K1(100) As String, arrROPDOSISRV_Code(100) As String, arrROPDOSISRV_Lib(100) As String
Dim arrROPDOSISRV_K As Integer, arrROPDOSISRV_ListIndex(100) As Integer
Dim arrROPDOSISRV_Mail(100) As String
Dim arrRecipient(50) As String, arrRecipient_Nb As Integer, blnRecipient As Boolean
Dim wRecipient As String, wccRecipient As String
'___________________________________________________________________________
Dim dupYROPDOS0 As typeYROPDOS0
Dim dupYROPINF0() As typeYROPINF0, dupYROPINF0_Nb As Long

Dim currentROPDOSISRV As String, currentROPDOSISRV_Nom As String
Dim currentROPDOSISRV_Hab As String, currentROPDOSISRV_Rôle As String
Dim mDateDiff_Duplication As Long
Dim blnYROPINF0_12X As Boolean, blnYROPINF0_12X_Aut As Boolean
Dim blnROPINFIDTL_Ok As Boolean
Dim blnROPINFIDT_Insérer As Boolean
Dim fraDossier_Left As Long, fraDossier_Right As Long

Dim blnSelect_Update_B_Display As Boolean
Dim blntxtUpdate_ROPINFGECH_Change As Boolean

Dim blnROPDOSQUAL As Boolean
Dim blnParam_ROPDOSQUAL As Boolean, blnParam_ROPDOSQUALB2 As Boolean
Dim mlstParam_ListIndex As Long
Dim arrROPDOSQUAL() As String, arrROPDOSQUAL_Code() As String, arrROPDOSQUAL_Nb As Long
Dim arrROPDOSQUALB2(10) As String

Dim blnParam_ROPDOSISRV As Boolean, kParam_ROPDOSISRV As Integer
'___________________________________________________________________________
Dim mDestinataire_Select As String, blnDestinataire_Select As Boolean
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim wDétail As String
Dim wTD_BackColor As String, wTD_ForeColor As String, wTD_Sta_ForeColor As String, wTD_Txt_ForeColor As String


Dim paramROPDOS_Path_DROPI As String
Dim blnUpdate_Sucess As Boolean
Dim mROPDOSIAMJ_Min As String

Dim currentYROPINF0 As typeYROPINF0
Dim arrROPINFMAIL() As String
Dim blnExportation_xlsx As Boolean
Dim blnIncidentSignificatif_Mail As Boolean

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim rsSabX As New ADODB.Recordset
Dim rsROPDOSGSRV As New ADODB.Recordset

Dim mXls2_Col As Integer, mXls2_Row As Integer

Dim blnROPDOSMAIL As Boolean, blnSécurité_Mail As Boolean
Dim oldROPDOSMAIL As typeYBIATAB0, newROPDOSMAIL As typeYBIATAB0
Dim mcboSelect_SQL_ListIndex As Integer
Public Sub YROPDOS0_Export()
On Error GoTo Error_Handler
Dim Nb As Long, wId As Long
Dim wFile As String, wFile2 As String, xSQL As String
Dim wAMJMin As String, WAMJMax As String
Dim X As String, K As Long, kMax As Long, K2 As Long, K3 As Long
Dim xFiltre As String, xROPDOSQUAL As String, mROPDOSQUAL As String
Dim arrS(10, 100) As Integer, kBale2 As Integer, kSrv As Integer
Dim wTotal As Long
Dim xPeriode As String
Dim arrDetail_Col(100) As Integer, arrDetail_Nb(100) As Integer
Dim wCount As Integer, totalCount As Integer
Dim xQual As String
Call DTPicker_Control(txtExport_AMJMIN, wAMJMin)
Call DTPicker_Control(txtExport_AMJMAX, WAMJMax)

xPeriode = "dossiers créés du " & dateImp10(wAMJMin) & " au " & dateImp10(WAMJMax)
wFile = "C:\temp\DROPI.xlsx"
If Dir(wFile) <> "" Then Kill wFile
wFile2 = "C:\temp\DROPI_B2.xlsx"
If Dir(wFile2) <> "" Then Kill wFile2

xFiltre = " and ROPDOSGPRV <> 'U' and ROPDOSSTA <> 'A'"
If chkExport_ROPDOSGNAT = "1" Then xFiltre = xFiltre & " and ROPDOSGNAT = 'I'"

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "RO détail des " & xPeriode
    .Subject = "RO dossiers"
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "RO Dossiers"
'__________________________________________________________________________________

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(0, 0, 255)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .WrapText = True
End With

Nb = 1
Call lstErr_AddItem(lstErr, cmdContext, "Export en cours : " & Nb & " enregistrements"): DoEvents

wsExcel.Cells(Nb, 1) = "ROPDOSID": wsExcel.Columns(1).ColumnWidth = 13: wsExcel.Columns(1).NumberFormat = "#######"
wsExcel.Cells(Nb, 2) = "ROPDOSSTA": wsExcel.Columns(2).ColumnWidth = 13
wsExcel.Cells(Nb, 3) = "ROPDOSSTAK": wsExcel.Columns(3).ColumnWidth = 13
wsExcel.Cells(Nb, 4) = "ROPDOSCUSR": wsExcel.Columns(4).ColumnWidth = 13
wsExcel.Cells(Nb, 5) = "ROPDOSCAMJ": wsExcel.Columns(5).ColumnWidth = 13: wsExcel.Columns(5).NumberFormat = "mm/dd/yyyy"
wsExcel.Columns(5).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
wsExcel.Cells(Nb, 6) = "ROPDOSUAMJ": wsExcel.Columns(6).ColumnWidth = 13: wsExcel.Columns(6).NumberFormat = "mm/dd/yyyy" '"hh:mm:ss"
wsExcel.Columns(6).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
wsExcel.Cells(Nb, 7) = "ROPDOSUVER": wsExcel.Columns(7).ColumnWidth = 13
wsExcel.Cells(Nb, 8) = "ROPDOSGECH": wsExcel.Columns(8).ColumnWidth = 13: wsExcel.Columns(8).NumberFormat = "mm/dd/yyyy"
wsExcel.Columns(8).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
wsExcel.Cells(Nb, 9) = "ROPDOSGUSR": wsExcel.Columns(9).ColumnWidth = 13
wsExcel.Cells(Nb, 10) = "ROPDOSGSRV": wsExcel.Columns(10).ColumnWidth = 13
wsExcel.Cells(Nb, 11) = "ROPDOSGNAT": wsExcel.Columns(11).ColumnWidth = 13
wsExcel.Cells(Nb, 12) = "ROPDOSGPRV": wsExcel.Columns(12).ColumnWidth = 13
wsExcel.Cells(Nb, 13) = "ROPDOSGGRA": wsExcel.Columns(13).ColumnWidth = 13
wsExcel.Cells(Nb, 14) = "ROPDOSGPRI": wsExcel.Columns(14).ColumnWidth = 13: wsExcel.Columns(14).NumberFormat = "#######"
wsExcel.Columns(14).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
wsExcel.Cells(Nb, 15) = "ROPDOSGCOU": wsExcel.Columns(15).ColumnWidth = 13
wsExcel.Cells(Nb, 16) = "ROPDOSIAMJ": wsExcel.Columns(16).ColumnWidth = 13: wsExcel.Columns(16).NumberFormat = "mm/dd/yyyy"
wsExcel.Columns(16).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
wsExcel.Cells(Nb, 17) = "ROPDOSISRV": wsExcel.Columns(17).ColumnWidth = 13
wsExcel.Cells(Nb, 18) = "ROPDOSIUSR": wsExcel.Columns(18).ColumnWidth = 13
wsExcel.Cells(Nb, 19) = "ROPDOSIREF": wsExcel.Columns(19).ColumnWidth = 20
wsExcel.Cells(Nb, 20) = "ROPDOSXDOM": wsExcel.Columns(20).ColumnWidth = 16
wsExcel.Cells(Nb, 21) = "ROPDOSXAPP": wsExcel.Columns(21).ColumnWidth = 16
wsExcel.Cells(Nb, 22) = "ROPDOSXID": wsExcel.Columns(22).ColumnWidth = 13
wsExcel.Cells(Nb, 23) = "ROPDOSQUAL": wsExcel.Columns(23).ColumnWidth = 13
wsExcel.Cells(Nb, 24) = "+ Bâle II": wsExcel.Columns(24).ColumnWidth = 20
wsExcel.Columns(24).Font.Size = 8
wsExcel.Cells(Nb, 25) = "+ qualification RO": wsExcel.Columns(25).ColumnWidth = 20
wsExcel.Columns(25).Font.Size = 8
wsExcel.Cells(Nb, 26) = "+ Service": wsExcel.Columns(26).ColumnWidth = 13
wsExcel.Cells(Nb, 27) = "+ Description": wsExcel.Columns(27).ColumnWidth = 200
wsExcel.Columns(27).Font.Size = 8

For K = 1 To 27
    wsExcel.Columns(K).VerticalAlignment = Excel.XlVAlign.xlVAlignTop
    wsExcel.Cells(1, K).Interior.ColorIndex = 8

Next K
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 " _
     & " where ROPDOSIAMJ >= '" & wAMJMin & "'" _
     & " and   ROPDOSIAMJ <= '" & WAMJMax & "'" _
     & xFiltre & " order by ROPDOSID"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    wId = rsSab("ROPDOSID")
    If wId >= 1000 And rsSab("ROPDOSGPRV") <> "U" Then
        Nb = Nb + 1
        wId = rsSab("ROPDOSID")
        wsExcel.Cells(Nb, 1) = wId
        wsExcel.Cells(Nb, 2) = rsSab("ROPDOSSTA")
        wsExcel.Cells(Nb, 3) = rsSab("ROPDOSSTAK")
        wsExcel.Cells(Nb, 4) = rsSab("ROPDOSCUSR")
        wsExcel.Cells(Nb, 5) = dateImp10_S(rsSab("ROPDOSCAMJ"))
        wsExcel.Cells(Nb, 6) = dateImp10_S(rsSab("ROPDOSUAMJ"))
        wsExcel.Cells(Nb, 7) = rsSab("ROPDOSUVER")
        wsExcel.Cells(Nb, 8) = dateImp10_S(rsSab("ROPDOSGECH"))
        wsExcel.Cells(Nb, 9) = rsSab("ROPDOSGUSR")
        wsExcel.Cells(Nb, 10) = rsSab("ROPDOSGSRV")
        wsExcel.Cells(Nb, 11) = rsSab("ROPDOSGNAT")
        wsExcel.Cells(Nb, 12) = rsSab("ROPDOSGPRV")
        wsExcel.Cells(Nb, 13) = rsSab("ROPDOSGGRA")
        wsExcel.Cells(Nb, 14) = Val(rsSab("ROPDOSGPRI"))
        wsExcel.Cells(Nb, 15) = rsSab("ROPDOSGCOU")
        wsExcel.Cells(Nb, 16) = dateImp10_S(rsSab("ROPDOSIAMJ"))
        wsExcel.Cells(Nb, 17) = rsSab("ROPDOSISRV")
        wsExcel.Cells(Nb, 18) = rsSab("ROPDOSIUSR")
        wsExcel.Cells(Nb, 19) = rsSab("ROPDOSIREF")
        wsExcel.Cells(Nb, 20) = rsSab("ROPDOSXDOM")
        wsExcel.Cells(Nb, 21) = rsSab("ROPDOSXAPP")
        wsExcel.Cells(Nb, 22) = rsSab("ROPDOSXID")
        xROPDOSQUAL = rsSab("ROPDOSQUAL")
        wsExcel.Cells(Nb, 23) = xROPDOSQUAL
        wsExcel.Cells(Nb, 24) = 9
        wsExcel.Cells(Nb, 25) = "????"
        For K = 1 To arrROPDOSQUAL_Nb
            If xROPDOSQUAL = arrROPDOSQUAL_Code(K) Then
                If Not IsNumeric(Mid$(arrROPDOSQUAL(K), 1, 1)) Then
                    kBale2 = 9
                Else
                    kBale2 = Mid$(arrROPDOSQUAL(K), 1, 1)
                End If
                
                wsExcel.Cells(Nb, 24) = arrROPDOSQUALB2(kBale2)
                wsExcel.Cells(Nb, 25) = arrROPDOSQUAL(K)
                Exit For
            End If
        Next K
        X = rsSab("ROPDOSISRV")
        If Mid$(X, 1, 2) = "_S" Then
            kSrv = Val(Mid$(X, 3, 2))
            wsExcel.Cells(Nb, 26) = arrROPDOSISRV_Code(kSrv)
        End If
                           
        If chkExport_ROPINFGTXT = "1" Then
            xSQL = "select ROPINFGTXT from " & paramIBM_Library_SABSPE & ".YROPINF0 " _
                 & " where ROPINFID = " & wId _
                 & " and   ROPINFIDP = 1 and ROPINFIDT =0  and ROPINFIDT2 = 1"
               
            Set rsSabX = cnsab.Execute(xSQL)
            If Not rsSabX.EOF Then
                wsExcel.Cells(Nb, 27) = rsSabX("ROPINFGTXT")
            End If
        End If
        arrS(kBale2, kSrv) = arrS(kBale2, kSrv) + 1
        If Nb Mod 10 = 0 Then Call lstErr_ChangeLastItem(lstErr, cmdContext, "Export en cours : " & Nb & " enregistrements"): DoEvents

    End If
    rsSab.MoveNext
Loop

Call lstErr_ChangeLastItem(lstErr, cmdContext, "Export en cours : " & Nb & " enregistrements"): DoEvents
Set rsSab = Nothing


wbExcel.SaveAs wFile

wbExcel.Close
'____________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Export stat Bâle2"): DoEvents
Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "RO Bâle II " & xPeriode
    .Subject = "Ventilation RO / service"
End With

Set wsExcel = wbExcel.ActiveSheet
With wsExcel
    .Cells.Font.Name = "Arial"
    .Name = "RO Bâle II"
    .PageSetup.CenterHorizontally = True
    .PageSetup.CenterHeader = "&U&B" & "&""Arial""" & "RO Bâle II " & xPeriode
    .PageSetup.LeftFooter = "&F" & " " & "&A"
    .PageSetup.RightFooter = "&D" & " " & "&T"
    .PageSetup.Orientation = xlLandscape
    .PageSetup.Zoom = False
    .PageSetup.FitToPagesTall = 1
    .PageSetup.FitToPagesWide = 1
    .PageSetup.PrintGridlines = True
End With
Nb = 1
wsExcel.Rows(1).RowHeight = 60
wsExcel.Cells(Nb, 1) = "Entités": wsExcel.Columns(1).ColumnWidth = 15
wsExcel.Cells(Nb, 2) = arrROPDOSQUALB2(1): wsExcel.Columns(2).ColumnWidth = 15: wsExcel.Columns(2).NumberFormat = "#######"
wsExcel.Cells(Nb, 3) = arrROPDOSQUALB2(2): wsExcel.Columns(3).ColumnWidth = 15: wsExcel.Columns(3).NumberFormat = "#######"
wsExcel.Cells(Nb, 4) = arrROPDOSQUALB2(3): wsExcel.Columns(4).ColumnWidth = 15: wsExcel.Columns(4).NumberFormat = "#######"
wsExcel.Cells(Nb, 5) = arrROPDOSQUALB2(4): wsExcel.Columns(5).ColumnWidth = 15: wsExcel.Columns(5).NumberFormat = "#######"
wsExcel.Cells(Nb, 6) = arrROPDOSQUALB2(5): wsExcel.Columns(6).ColumnWidth = 15: wsExcel.Columns(6).NumberFormat = "#######"
wsExcel.Cells(Nb, 7) = arrROPDOSQUALB2(6): wsExcel.Columns(7).ColumnWidth = 15: wsExcel.Columns(7).NumberFormat = "#######"
wsExcel.Cells(Nb, 8) = arrROPDOSQUALB2(7): wsExcel.Columns(8).ColumnWidth = 15: wsExcel.Columns(8).NumberFormat = "#######"
wsExcel.Cells(Nb, 9) = arrROPDOSQUALB2(8): wsExcel.Columns(9).ColumnWidth = 15: wsExcel.Columns(9).NumberFormat = "#######"
wsExcel.Cells(Nb, 10) = arrROPDOSQUALB2(9): wsExcel.Columns(10).ColumnWidth = 15: wsExcel.Columns(10).NumberFormat = "#######"
wsExcel.Cells(Nb, 11) = "Total": wsExcel.Columns(11).ColumnWidth = 15: wsExcel.Columns(11).NumberFormat = "#######"

arrROPDOSISRV_Code(100) = "Total"
For kSrv = 1 To 100
    If Mid$(arrROPDOSISRV_Code(kSrv), 1, 1) <> "?" Then
        Nb = Nb + 1
        wTotal = 0
        wsExcel.Cells(Nb, 1) = arrROPDOSISRV_Code(kSrv)
        For kBale2 = 1 To 10
            wsExcel.Cells(Nb, kBale2 + 1) = arrS(kBale2, kSrv)
            If kBale2 < 10 Then
                arrS(10, kSrv) = arrS(10, kSrv) + arrS(kBale2, kSrv)
                If kSrv < 100 Then
                    arrS(kBale2, 100) = arrS(kBale2, 100) + arrS(kBale2, kSrv)
                End If
           End If
        Next kBale2
    End If
Next kSrv

For K = 1 To 11
   If K > 1 Then wsExcel.Columns(K).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
   With wsExcel.Cells(1, K)
        .HorizontalAlignment = Excel.xlHAlignCenter
        .VerticalAlignment = Excel.xlVAlignCenter
        .WrapText = True
        .Interior.Color = RGB(255, 255, 128)
        .Font.Bold = True
    End With
                 
    wsExcel.Cells(Nb, K).Interior.Color = RGB(255, 255, 128)
    wsExcel.Cells(Nb, K).Font.Bold = True

Next K
For K = 1 To Nb
    wsExcel.Cells(K, 11).Interior.Color = RGB(255, 255, 128)
    wsExcel.Cells(K, 11).Font.Bold = True
    If K < 1 Then wsExcel.Rows(K).RowHeight = 18

Next K
With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(0, 0, 255)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
   
End With
'____________________________________________________________________________________

'wbExcel.Worksheets.Add
Set wsExcel = wbExcel.Sheets(2) '   .ActiveSheet
With wsExcel
    .Cells.Font.Name = "Arial"
    .Name = "RO Bâle II Détail"
    .PageSetup.CenterHorizontally = True
    .PageSetup.CenterHeader = "&U&B" & "&""Arial""" & "RO Bâle II Détail" & xPeriode
    .PageSetup.LeftFooter = "&F" & " " & "&A"
    .PageSetup.RightFooter = "&D" & " " & "&T"
    .PageSetup.Orientation = xlLandscape
    .PageSetup.Zoom = False
    .PageSetup.FitToPagesTall = 1
    .PageSetup.FitToPagesWide = 1
    .PageSetup.PrintGridlines = True
End With
With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(0, 0, 255)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .WrapText = True
End With
With wsExcel.Columns(3)
    .HorizontalAlignment = Excel.xlHAlignLeft
End With

Nb = 1
wsExcel.Cells(Nb, 1) = "B2": wsExcel.Columns(1).ColumnWidth = 3
wsExcel.Cells(Nb, 2) = "Qual": wsExcel.Columns(2).ColumnWidth = 5
wsExcel.Cells(Nb, 3) = "Libellé": wsExcel.Columns(3).ColumnWidth = 25
kMax = 3
For kSrv = 1 To 100
    If Mid$(arrROPDOSISRV_Code(kSrv), 1, 1) <> "?" Then
        kMax = kMax + 1
        wsExcel.Cells(Nb, kMax) = arrROPDOSISRV_Code(kSrv)
        arrDetail_Col(kSrv) = kMax

    End If
Next kSrv

xSQL = "select count(*) ,ROPDOSQUAL , ROPDOSISRV  from " & paramIBM_Library_SABSPE & ".YROPDOS0 " _
     & " where ROPDOSIAMJ >= '" & wAMJMin & "'" _
     & " and   ROPDOSIAMJ <= '" & WAMJMax & "'" _
     & xFiltre & " group by ROPDOSQUAL , ROPDOSISRV order by ROPDOSQUAL , ROPDOSISRV"
Set rsSab = cnsab.Execute(xSQL)

mROPDOSQUAL = ""
totalCount = 0
Do While Not rsSab.EOF
    xROPDOSQUAL = rsSab("ROPDOSQUAL")
    If mROPDOSQUAL <> xROPDOSQUAL Then
        If totalCount > 0 Then wsExcel.Cells(Nb, arrDetail_Col(100)) = totalCount
        totalCount = 0
        Nb = Nb + 1
        mROPDOSQUAL = xROPDOSQUAL
        wsExcel.Cells(Nb, 2) = mROPDOSQUAL
        For K2 = 1 To arrROPDOSQUAL_Nb
            If mROPDOSQUAL = arrROPDOSQUAL_Code(K2) Then
                xQual = arrROPDOSQUAL(K2)
                If IsNumeric(Mid$(xQual, 1, 1)) Then
                    wsExcel.Cells(Nb, 1) = Mid$(xQual, 1, 1)
                    wsExcel.Cells(Nb, 3) = Mid$(xQual, 3, Len(xQual) - 2)
               Else
                    wsExcel.Cells(Nb, 1) = 9
                    wsExcel.Cells(Nb, 3) = xQual
               End If
                Exit For
            End If
        Next K2
    End If
    kSrv = Val(Mid$(rsSab("ROPDOSISRV"), 3, 2))
    wCount = rsSab(0)
    totalCount = totalCount + wCount
    wsExcel.Cells(Nb, arrDetail_Col(kSrv)) = wCount
    arrDetail_Nb(kSrv) = arrDetail_Nb(kSrv) + wCount
    rsSab.MoveNext
Loop
wsExcel.Cells(Nb, arrDetail_Col(100)) = totalCount

Nb = Nb + 1
wTotal = 0
For kSrv = 1 To 100
    If arrDetail_Col(kSrv) <> 0 And arrDetail_Nb(kSrv) <> 0 Then
        wsExcel.Cells(Nb, arrDetail_Col(kSrv)) = arrDetail_Nb(kSrv)
        wTotal = wTotal + arrDetail_Nb(kSrv)
    End If
Next kSrv
wsExcel.Cells(Nb, kMax) = wTotal
wsExcel.Cells(Nb, 3) = "Total"
        
For K = 1 To kMax
    If K > 3 Then wsExcel.Columns(K).ColumnWidth = 6.5
    wsExcel.Cells(1, K).Interior.Color = RGB(255, 200, 100)
    wsExcel.Cells(Nb, K).Interior.Color = RGB(255, 200, 100)
Next K
For K = 1 To Nb
    wsExcel.Rows(K).RowHeight = 27
    wsExcel.Cells(K, kMax).Interior.Color = RGB(255, 200, 100)
    wsExcel.Cells(K, 1).Interior.Color = RGB(255, 200, 100)
    wsExcel.Cells(K, 2).Interior.Color = RGB(255, 200, 100)
Next K
 

Set rsSab = Nothing

wbExcel.SaveAs wFile2

wbExcel.Close

'____________________________________________________________________________________
appExcel.Quit

Set rsSab = Nothing
Set rsSabX = Nothing

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
Call lstErr_AddItem(lstErr, cmdContext, "Export terminé : " & Nb & " enregistrements"): DoEvents

'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub

Public Sub YROPDOS0_Xls_Migration()
On Error GoTo Error_Handler
Dim Nb As Long, wId As Long
Dim wFile As String, wFile2 As String, xSQL As String
Dim wAMJMin As String, WAMJMax As String
Dim X As String, K As Long, kMax As Long, K2 As Long, K3 As Long
Dim xFiltre As String, xROPDOSQUAL As String, mROPDOSQUAL As String
Dim arrS(10, 100) As Integer, kBale2 As Integer, kSrv As Integer
Dim wTotal As Long
Dim xPeriode As String
Dim arrDetail_Col(100) As Integer, arrDetail_Nb(100) As Integer
Dim wCount As Integer, totalCount As Integer
Dim xQual As String
Call DTPicker_Control(txtExport_AMJMIN, wAMJMin)
Call DTPicker_Control(txtExport_AMJMAX, WAMJMax)

xPeriode = "dossiers créés du " & dateImp10(wAMJMin) & " au " & dateImp10(WAMJMax)
wFile = "C:\temp\DROPI.xlsx"
If Dir(wFile) <> "" Then Kill wFile

'xFiltre = " where ROPDOSIAMJ >= '" & wAMJMin & "'" _
'        & " and   ROPDOSIAMJ <= '" & WAMJMax & "'" _
'        & " and ROPDOSID > 1000  and ROPDOSGNAT = 'I' and ROPDOSGPRV <> 'U'"

xFiltre = " where ROPDOSID > 1000  and ROPDOSGNAT = 'I' and ROPDOSGPRV <> 'U'"

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "DROPI" & xPeriode
    .Subject = "DROPI"
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "Dossiers"
'__________________________________________________________________________________

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(0, 0, 255)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .WrapText = True
End With

Nb = 1
Call lstErr_AddItem(lstErr, cmdContext, "Export en cours : " & Nb & " enregistrements"): DoEvents

wsExcel.Cells(Nb, 1) = "ROPDOSID": wsExcel.Columns(1).ColumnWidth = 13: wsExcel.Columns(1).NumberFormat = "#######"
wsExcel.Cells(Nb, 2) = "ROPDOSSTA": wsExcel.Columns(2).ColumnWidth = 13
wsExcel.Cells(Nb, 3) = "ROPDOSCUSR": wsExcel.Columns(3).ColumnWidth = 13
wsExcel.Cells(Nb, 4) = "ROPDOSCAMJ": wsExcel.Columns(4).ColumnWidth = 13
wsExcel.Cells(Nb, 5) = "ROPDOSUAMJ": wsExcel.Columns(5).ColumnWidth = 13
wsExcel.Cells(Nb, 6) = "ROPDOSUHMS": wsExcel.Columns(6).ColumnWidth = 13
wsExcel.Cells(Nb, 7) = "ROPDOSGECH": wsExcel.Columns(7).ColumnWidth = 13
wsExcel.Cells(Nb, 8) = "ROPDOSGUSR": wsExcel.Columns(8).ColumnWidth = 13
wsExcel.Cells(Nb, 9) = "ROPDOSGSRV": wsExcel.Columns(9).ColumnWidth = 13
wsExcel.Cells(Nb, 10) = "ROPDOSGNAT": wsExcel.Columns(10).ColumnWidth = 13
wsExcel.Cells(Nb, 11) = "ROPDOSGPRV": wsExcel.Columns(11).ColumnWidth = 13
wsExcel.Cells(Nb, 12) = "ROPDOSGGRA": wsExcel.Columns(12).ColumnWidth = 13
wsExcel.Cells(Nb, 13) = "ROPDOSGPRI": wsExcel.Columns(13).ColumnWidth = 13: wsExcel.Columns(13).NumberFormat = "#######"
wsExcel.Cells(Nb, 14) = "ROPDOSGCOU": wsExcel.Columns(14).ColumnWidth = 13
wsExcel.Cells(Nb, 15) = "ROPDOSIAMJ": wsExcel.Columns(15).ColumnWidth = 13
wsExcel.Cells(Nb, 16) = "ROPDOSISRV": wsExcel.Columns(16).ColumnWidth = 13
wsExcel.Cells(Nb, 17) = "ROPDOSIUSR": wsExcel.Columns(17).ColumnWidth = 13
wsExcel.Cells(Nb, 18) = "ROPDOSIREF": wsExcel.Columns(18).ColumnWidth = 20
wsExcel.Cells(Nb, 19) = "ROPDOSXDOM": wsExcel.Columns(19).ColumnWidth = 16
wsExcel.Cells(Nb, 20) = "ROPDOSXAPP": wsExcel.Columns(20).ColumnWidth = 16
wsExcel.Cells(Nb, 21) = "ROPDOSXID": wsExcel.Columns(21).ColumnWidth = 13
wsExcel.Cells(Nb, 22) = "ROPDOSQUAL": wsExcel.Columns(22).ColumnWidth = 13

For K = 1 To 22
    wsExcel.Columns(K).VerticalAlignment = Excel.XlVAlign.xlVAlignTop
    wsExcel.Cells(1, K).Interior.ColorIndex = 8

Next K
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 " _
     & xFiltre & " order by ROPDOSID"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    wId = rsSab("ROPDOSID")
    If wId >= 1000 And rsSab("ROPDOSGPRV") <> "U" Then
        Nb = Nb + 1
        wId = rsSab("ROPDOSID")
        wsExcel.Cells(Nb, 1) = wId
        wsExcel.Cells(Nb, 2) = rsSab("ROPDOSSTA")
        wsExcel.Cells(Nb, 3) = rsSab("ROPDOSCUSR")
        wsExcel.Cells(Nb, 4) = rsSab("ROPDOSCAMJ")
        wsExcel.Cells(Nb, 5) = rsSab("ROPDOSUAMJ")
        wsExcel.Cells(Nb, 6) = rsSab("ROPDOSUHMS")
        wsExcel.Cells(Nb, 7) = rsSab("ROPDOSGECH")
        'wsExcel.Cells(Nb, 8) = rsSab("ROPDOSGUSR")
        'wsExcel.Cells(Nb, 9) = rsSab("ROPDOSGSRV")
        wsExcel.Cells(Nb, 10) = rsSab("ROPDOSGNAT")
        wsExcel.Cells(Nb, 11) = rsSab("ROPDOSGPRV")
        wsExcel.Cells(Nb, 12) = rsSab("ROPDOSGGRA")
        wsExcel.Cells(Nb, 13) = Val(rsSab("ROPDOSGPRI"))
        wsExcel.Cells(Nb, 14) = rsSab("ROPDOSGCOU")
        wsExcel.Cells(Nb, 15) = rsSab("ROPDOSIAMJ")
        'wsExcel.Cells(Nb, 16) = rsSab("ROPDOSISRV")
        'wsExcel.Cells(Nb, 17) = rsSab("ROPDOSIUSR")
        wsExcel.Cells(Nb, 18) = rsSab("ROPDOSIREF")
        wsExcel.Cells(Nb, 19) = rsSab("ROPDOSXDOM")
        wsExcel.Cells(Nb, 20) = rsSab("ROPDOSXAPP")
        wsExcel.Cells(Nb, 21) = rsSab("ROPDOSXID")
        xROPDOSQUAL = rsSab("ROPDOSQUAL")
        wsExcel.Cells(Nb, 22) = xROPDOSQUAL
        
        X = rsSab("ROPDOSGUSR")
        If Mid$(X, 1, 2) = "_S" Then
            kSrv = Val(Mid$(X, 3, 2))
            wsExcel.Cells(Nb, 8) = arrROPDOSISRV_Code(kSrv)
        Else
            wsExcel.Cells(Nb, 8) = X
        End If
        X = rsSab("ROPDOSGSRV")
        If Mid$(X, 1, 2) = "_S" Then
            kSrv = Val(Mid$(X, 3, 2))
            wsExcel.Cells(Nb, 9) = arrROPDOSISRV_Code(kSrv)
        Else
            wsExcel.Cells(Nb, 9) = X
        End If
         
        X = rsSab("ROPDOSISRV")
        If Mid$(X, 1, 2) = "_S" Then
            kSrv = Val(Mid$(X, 3, 2))
            wsExcel.Cells(Nb, 16) = arrROPDOSISRV_Code(kSrv)
        Else
            wsExcel.Cells(Nb, 16) = X
        End If
        X = rsSab("ROPDOSIUSR")
        If Mid$(X, 1, 2) = "_S" Then
            kSrv = Val(Mid$(X, 3, 2))
            wsExcel.Cells(Nb, 17) = arrROPDOSISRV_Code(kSrv)
        Else
            wsExcel.Cells(Nb, 17) = X
        End If
                          
        If Nb Mod 10 = 0 Then Call lstErr_ChangeLastItem(lstErr, cmdContext, "Export Dossiers : " & Nb & " enregistrements"): DoEvents

    End If
    rsSab.MoveNext
Loop

Call lstErr_ChangeLastItem(lstErr, cmdContext, "Export en cours : " & Nb & " enregistrements"): DoEvents

'=========================================================================================================
Set wsExcel = wbExcel.Sheets(2)
wsExcel.Name = "Informations"
'__________________________________________________________________________________

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(0, 0, 255)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .WrapText = True
End With

Nb = 1
Call lstErr_AddItem(lstErr, cmdContext, "Export en cours : " & Nb & " enregistrements"): DoEvents

wsExcel.Cells(Nb, 1) = "ROPINFID": wsExcel.Columns(1).ColumnWidth = 13: wsExcel.Columns(1).NumberFormat = "######0"
wsExcel.Cells(Nb, 2) = "ROPINFIDP": wsExcel.Columns(2).ColumnWidth = 13: wsExcel.Columns(2).NumberFormat = "######0"
wsExcel.Cells(Nb, 3) = "ROPINFIDT": wsExcel.Columns(3).ColumnWidth = 13: wsExcel.Columns(3).NumberFormat = "######0"
wsExcel.Cells(Nb, 4) = "ROPINFIDT2": wsExcel.Columns(4).ColumnWidth = 13: wsExcel.Columns(4).NumberFormat = "######0"
wsExcel.Cells(Nb, 5) = "ROPINFSTA": wsExcel.Columns(5).ColumnWidth = 13
wsExcel.Cells(Nb, 6) = "ROPINFCUSR": wsExcel.Columns(6).ColumnWidth = 13
wsExcel.Cells(Nb, 7) = "ROPINFCAMJ": wsExcel.Columns(7).ColumnWidth = 13
wsExcel.Cells(Nb, 8) = "ROPINFUUSR": wsExcel.Columns(8).ColumnWidth = 13
wsExcel.Cells(Nb, 9) = "ROPINFUAMJ": wsExcel.Columns(9).ColumnWidth = 13
wsExcel.Cells(Nb, 10) = "ROPINFUHMS": wsExcel.Columns(10).ColumnWidth = 13
wsExcel.Cells(Nb, 11) = "ROPINFGECH": wsExcel.Columns(11).ColumnWidth = 13
wsExcel.Cells(Nb, 12) = "ROPINFGUSR": wsExcel.Columns(12).ColumnWidth = 13
wsExcel.Cells(Nb, 13) = "ROPINFGSRV": wsExcel.Columns(13).ColumnWidth = 13
wsExcel.Cells(Nb, 14) = "ROPINFGNAT": wsExcel.Columns(14).ColumnWidth = 13
wsExcel.Cells(Nb, 15) = "ROPINFGTXT": wsExcel.Columns(15).ColumnWidth = 100

For K = 1 To 15
    wsExcel.Columns(K).VerticalAlignment = Excel.XlVAlign.xlVAlignTop
    wsExcel.Cells(1, K).Interior.ColorIndex = 8

Next K
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YROPINF0 " _
     & " where ROPINFID in (select ROPDOSID from " & paramIBM_Library_SABSPE & ".YROPDOS0 " _
     & xFiltre & ") order by ROPINFID , ROPINFIDP , ROPINFIDT , ROPINFIDT2"

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    wId = rsSab("ROPINFID")
        Nb = Nb + 1
        wId = rsSab("ROPINFID")
        wsExcel.Cells(Nb, 1) = wId
        wsExcel.Cells(Nb, 2) = rsSab("ROPINFIDP")
        wsExcel.Cells(Nb, 3) = rsSab("ROPINFIDT")
        wsExcel.Cells(Nb, 4) = rsSab("ROPINFIDT2")
        wsExcel.Cells(Nb, 5) = rsSab("ROPINFSTA")
        wsExcel.Cells(Nb, 6) = rsSab("ROPINFCUSR")
        wsExcel.Cells(Nb, 7) = rsSab("ROPINFCAMJ")
        wsExcel.Cells(Nb, 8) = rsSab("ROPINFUUSR")
        wsExcel.Cells(Nb, 9) = rsSab("ROPINFUAMJ")
        wsExcel.Cells(Nb, 10) = rsSab("ROPINFUHMS")
        wsExcel.Cells(Nb, 11) = rsSab("ROPINFGECH")
        'wsExcel.Cells(Nb, 12) = rsSab("ROPINFGUSR")
        'wsExcel.Cells(Nb, 13) = rsSab("ROPINFGSRV")
        wsExcel.Cells(Nb, 14) = rsSab("ROPINFGNAT")
        wsExcel.Cells(Nb, 15) = rsSab("ROPINFGTXT")
        
        X = rsSab("ROPINFGUSR")
        If Mid$(X, 1, 2) = "_S" Then
            kSrv = Val(Mid$(X, 3, 2))
            wsExcel.Cells(Nb, 12) = arrROPDOSISRV_Code(kSrv)
        Else
            wsExcel.Cells(Nb, 12) = X
        End If

        
        X = rsSab("ROPINFGSRV")
        If Mid$(X, 1, 2) = "_S" Then
            kSrv = Val(Mid$(X, 3, 2))
            wsExcel.Cells(Nb, 13) = arrROPDOSISRV_Code(kSrv)
        Else
            wsExcel.Cells(Nb, 13) = X
        End If
        
                          
        If Nb Mod 10 = 0 Then Call lstErr_ChangeLastItem(lstErr, cmdContext, "Export Info : " & Nb & " enregistrements"): DoEvents

    rsSab.MoveNext
Loop
'=========================================================================================================


Set rsSab = Nothing


wbExcel.SaveAs wFile

wbExcel.Close
'____________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Export stat Bâle2"): DoEvents

'____________________________________________________________________________________
appExcel.Quit

Set rsSab = Nothing
Set rsSabX = Nothing

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
Call lstErr_AddItem(lstErr, cmdContext, "Export terminé : " & Nb & " enregistrements"): DoEvents

'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub

Public Sub YROPDOS0_Xls_Service()
On Error GoTo Error_Handler
Dim X As String, K As Long, xWhere As String
Dim wFile As String, wFilex As String
Dim blnCALCS As Boolean
Dim xSQL As String
Dim wROPDOSGSRV As String, wROPDOSGSRV_Lib As String
On Error GoTo Error_Handler
'===================================================================================
ReDim arrYROPINF0(3)


'_________________________________________

xSQL = "select distinct ROPDOSGSRV from " & paramIBM_Library_SABSPE & ".YROPDOS0 " _
       & " where ROPDOSGPRV <> 'U' and ROPDOSSTA = ' ' and ROPDOSID > 1000" _
       & " and ROPDOSGNAT = 'I' order by ROPDOSGSRV"
       
Set rsROPDOSGSRV = cnsab.Execute(xSQL)

Do While Not rsROPDOSGSRV.EOF

    wROPDOSGSRV = rsROPDOSGSRV("ROPDOSGSRV")
    If Mid$(wROPDOSGSRV, 1, 2) = "_S" Then
        wROPDOSGSRV_Lib = Trim(arrROPDOSISRV_Code(Val(Mid$(wROPDOSGSRV, 3, 2))))
    Else
        wROPDOSGSRV_Lib = Trim(wROPDOSGSRV)
    End If

    wFile = "C:\Temp\DROPI" & " " & dateImp_Amj(DSys) & " " & wROPDOSGSRV_Lib & ".xlsx"
    If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile
    
    Set appExcel = CreateObject("Excel.Application")
    appExcel.Workbooks.Add
    Set wbExcel = appExcel.ActiveWorkbook
    With wbExcel
        .Title = "DROPI"
        .Subject = ""
    End With
    
    Call YROPDOS0_Xls_Service_Detail(wROPDOSGSRV, wROPDOSGSRV_Lib)
    wbExcel.SaveAs wFile
    wbExcel.Close
    appExcel.Quit
'===================================================================================================
    rsROPDOSGSRV.MoveNext
Loop

'_________________________________________



'__________________________________________________________________________________
Exit_sub:
'__________________________________________________________________________________

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

End Sub


Public Sub YROPDOS0_Xls_Service_Detail(lROPDOSGSRV As String, lROPDOSGSRV_Lib As String)
On Error GoTo Error_Handler
Dim X As String
Dim K As Integer, K1 As Integer
Dim IDP As Integer
Dim blnOk As Boolean

Dim wColor As Long

Call rsYROPDOS0_Init(oldYROPDOS0)


'==========================================================================================================

Set wsExcel = wbExcel.Sheets(1)
wsExcel.Name = "Dossiers"

'__________________________________________________________________________________

With wsExcel.Cells
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(160, 160, 160)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(220, 220, 220)
    .VerticalAlignment = Excel.xlVAlignCenter
    .HorizontalAlignment = Excel.xlHAlignLeft
    .WrapText = True
    .Font.Size = 8
    .Font.Name = "Arial Unicode MS"
    .Font.Color = RGB(0, 64, 128)
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14DROPI " & lROPDOSGSRV_Lib _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True
wsExcel.PageSetup.PrintTitleRows = "$A1:$H1"
wsExcel.PageSetup.Zoom = 75


Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : " & lROPDOSGSRV & " " & lROPDOSGSRV_Lib): DoEvents

wsExcel.Columns(1).ColumnWidth = 7: wsExcel.Cells(1, 1) = "Dossier": wsExcel.Columns(1).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
wsExcel.Columns(1).Font.Bold = True
wsExcel.Columns(2).ColumnWidth = 12: wsExcel.Cells(1, 2) = "Domaine"
wsExcel.Columns(3).ColumnWidth = 12: wsExcel.Cells(1, 3) = "Application"
wsExcel.Columns(4).ColumnWidth = 15: wsExcel.Cells(1, 4) = "Réf. externe"
wsExcel.Columns(5).ColumnWidth = 10: wsExcel.Cells(1, 5) = "D.création": wsExcel.Columns(5).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
wsExcel.Columns(6).ColumnWidth = 12: wsExcel.Cells(1, 6) = "Srv initiateur"
wsExcel.Columns(7).ColumnWidth = 60: wsExcel.Cells(1, 7) = "Description"
wsExcel.Columns(8).ColumnWidth = 40: wsExcel.Cells(1, 8) = "Commentaire"
wsExcel.Columns(9).ColumnWidth = 0
wsExcel.Columns(10).ColumnWidth = 0

wsExcel.Cells.EntireRow.AutoFit
mXls2_Col = 8
For K = 1 To mXls2_Col
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next


'==========================================================================================================
X = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 " _
       & " where ROPDOSGPRV <> 'U' and ROPDOSSTA = ' ' and ROPDOSID > 1000" _
       & "  and ROPDOSGNAT = 'I' and ROPDOSGSRV = '" & lROPDOSGSRV & "'" _
       & " order by ROPDOSID"
Set rsSab = cnsab.Execute(X)
mXls2_Row = 1
Do While Not rsSab.EOF
    V = rsYROPDOS0_GetBuffer(rsSab, xYROPDOS0)
    
    'blnOk = True
    mXls2_Row = mXls2_Row + 1
    wsExcel.Cells(mXls2_Row, 1) = xYROPDOS0.ROPDOSID
    wsExcel.Cells(mXls2_Row, 2) = xYROPDOS0.ROPDOSXDOM
    wsExcel.Cells(mXls2_Row, 3) = xYROPDOS0.ROPDOSXAPP
    wsExcel.Cells(mXls2_Row, 4) = xYROPDOS0.ROPDOSXID
    wsExcel.Cells(mXls2_Row, 5) = dateImp10(xYROPDOS0.ROPDOSCAMJ)
    If Mid$(xYROPDOS0.ROPDOSISRV, 1, 2) = "_S" Then
        wsExcel.Cells(mXls2_Row, 6) = arrROPDOSISRV_Code(Val(Mid$(xYROPDOS0.ROPDOSISRV, 3, 2)))
    Else
        wsExcel.Cells(mXls2_Row, 6) = xYROPDOS0.ROPDOSISRV
    End If
'_______________________________________________________________________
    arrYROPINF0_Nb = 0: IDP = 0
    X = "select * from " & paramIBM_Library_SABSPE & ".YROPINF0 " _
         & " where ROPINFID = " & xYROPDOS0.ROPDOSID _
         & " order by ropinfidp desc, ropinfidt desc, ropinfidt2 desc"
         '& " and   ROPINFIDP = 1 and ROPINFIDT =0  and ROPINFIDT2 = 1"
    Set rsSabX = cnsab.Execute(X)
    Do While Not rsSabX.EOF
        xYROPINF0.ROPINFGNAT = rsSabX("ROPINFGNAT")
        Select Case xYROPINF0.ROPINFGNAT
            Case "P":
                        wsExcel.Cells(mXls2_Row + IDP, 7) = vbCrLf & Trim(rsSabX("ROPINFGTXT")) & vbCrLf
                        wsExcel.Cells(mXls2_Row + IDP, 8).Interior.Color = RGB(255, 255, 220)
                        IDP = IDP + 1
            Case "A", "F", "N":
                If arrYROPINF0_Nb < 2 Then
                   arrYROPINF0_Nb = arrYROPINF0_Nb + 1
                   V = rsYROPINF0_GetBuffer(rsSabX, arrYROPINF0(arrYROPINF0_Nb))
                End If
                
       End Select
        rsSabX.MoveNext
    Loop
    
    For K1 = arrYROPINF0_Nb To 1 Step -1
        wsExcel.Cells(mXls2_Row + IDP, 7) = Replace(Trim(arrYROPINF0(K1).ROPINFGTXT), "=", " ")
        wsExcel.Cells(mXls2_Row + IDP, 7).Font.Color = RGB(128, 0, 128)
        If Trim(arrYROPINF0(K1).ROPINFCUSR) <> "" Then
            wsExcel.Cells(mXls2_Row + IDP, 5) = dateImp10(arrYROPINF0(K1).ROPINFCAMJ)
            wsExcel.Cells(mXls2_Row + IDP, 6) = arrYROPINF0(K1).ROPINFCUSR
        Else
            wsExcel.Cells(mXls2_Row + IDP, 5) = dateImp10(arrYROPINF0(K1).ROPINFUAMJ)
            wsExcel.Cells(mXls2_Row + IDP, 6) = arrYROPINF0(K1).ROPINFUUSR
        End If
        wsExcel.Cells(mXls2_Row + IDP, 5).Font.Color = RGB(128, 0, 128)
        wsExcel.Cells(mXls2_Row + IDP, 6).Font.Color = RGB(128, 0, 128)
       
        IDP = IDP + 1
    Next K1
    If IDP > 0 Then mXls2_Row = mXls2_Row + IDP - 1
    wsExcel.Range("A" & mXls2_Row & ":H" & mXls2_Row).Borders(xlEdgeBottom).Weight = xlMedium
    wsExcel.Range("A" & mXls2_Row & ":H" & mXls2_Row).Borders(xlEdgeBottom).Color = RGB(160, 160, 160)
   
'===================================================================================================
    rsSab.MoveNext
Loop
wsExcel.Range("A1:H1").Borders(xlEdgeTop).Weight = xlMedium
wsExcel.Range("A1:H1").Borders(xlEdgeTop).Color = RGB(96, 96, 96)
wsExcel.Range("A" & mXls2_Row & ":H" & mXls2_Row).Borders(xlEdgeBottom).Weight = xlMedium
wsExcel.Range("A" & mXls2_Row & ":H" & mXls2_Row).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)

wsExcel.Range("A1:A" & mXls2_Row).Borders(xlEdgeLeft).Weight = xlMedium
wsExcel.Range("A1:A" & mXls2_Row).Borders(xlEdgeLeft).Color = RGB(96, 96, 96)
wsExcel.Range("H1:H" & mXls2_Row).Borders(xlEdgeRight).Weight = xlMedium
wsExcel.Range("H1:H" & mXls2_Row).Borders(xlEdgeRight).Color = RGB(128, 128, 128)

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name


End Sub

Public Sub lstParam_Export()
On Error GoTo Error_Handler
Dim Nb As Long, wId As Long
Dim wFile As String, wFile2 As String, xSQL As String
Dim wAMJMin As String, WAMJMax As String
Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim X As String, K As Long, kMax As Long, K2 As Long, K3 As Long
Dim xFiltre As String, xROPDOSQUAL As String, mROPDOSQUAL As String
Dim rsSabX As New ADODB.Recordset
Dim arrDetail_Col(100) As Integer, kSrv As Integer
Dim X1 As String
'____________________________________________________________________________________

wFile2 = "C:\temp\DROPI_Param.xlsx"
If Dir(wFile2) <> "" Then Kill wFile2

Call lstErr_AddItem(lstErr, cmdContext, "Export param Services"): DoEvents

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "RO Services"
    .Subject = "services"
End With

Set wsExcel = wbExcel.ActiveSheet
With wsExcel
    .Cells.Font.Name = "Arial"
    .Name = "RO_services"
    '.PageSetup.CenterHorizontally = True
    .PageSetup.CenterHeader = "&U&B" & "&""Arial""" & "RO services "
    .PageSetup.LeftFooter = "&F" & " " & "&A"
    .PageSetup.RightFooter = "&D" & " " & "&T"
    .PageSetup.Orientation = xlLandscape
    .PageSetup.Zoom = False
    .PageSetup.FitToPagesTall = 1
    .PageSetup.FitToPagesWide = 1
    .PageSetup.PrintGridlines = True
End With
Nb = 1
wsExcel.Rows(1).RowHeight = 15
wsExcel.Cells(Nb, 1) = "Code": wsExcel.Columns(1).ColumnWidth = 10
wsExcel.Cells(Nb, 2) = "Sigle": wsExcel.Columns(2).ColumnWidth = 17
wsExcel.Cells(Nb, 3) = "Libellé": wsExcel.Columns(3).ColumnWidth = 40
For K = 1 To 3
    wsExcel.Cells(1, K).Interior.Color = RGB(255, 200, 100)
Next K

For kSrv = 1 To 100
    If Mid$(arrROPDOSISRV_Code(kSrv), 1, 1) <> "?" Then
        Nb = Nb + 1
        wsExcel.Cells(Nb, 1) = "S" & Format(kSrv, "00")
        wsExcel.Cells(Nb, 2) = Trim(arrROPDOSISRV_Code(kSrv))
        wsExcel.Cells(Nb, 3) = Trim(arrROPDOSISRV_Lib(kSrv))
    End If
Next kSrv

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(0, 0, 255)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
   
End With
'____________________________________________________________________________________

'wbExcel.Worksheets.Add
Set wsExcel = wbExcel.Sheets(2) '   .ActiveSheet
With wsExcel
    .Cells.Font.Name = "Arial"
    .Name = "RO_Utilisateurs"
    .PageSetup.CenterHorizontally = True
    .PageSetup.CenterHeader = "&U&B" & "&""Arial""" & "RO : Utilisateurs / services"
    .PageSetup.LeftFooter = "&F" & " " & "&A"
    .PageSetup.RightFooter = "&D" & " " & "&T"
    .PageSetup.Orientation = xlLandscape
    .PageSetup.Zoom = False
    .PageSetup.FitToPagesTall = 1
    .PageSetup.FitToPagesWide = 1
    .PageSetup.PrintGridlines = True
End With
With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(0, 0, 255)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .WrapText = True
End With

Nb = 1
wsExcel.Rows(1).RowHeight = 30

wsExcel.Cells(Nb, 1) = "Utilisateurs": wsExcel.Columns(1).ColumnWidth = 17
wsExcel.Cells(Nb, 2) = "Code": wsExcel.Columns(2).ColumnWidth = 7
wsExcel.Cells(Nb, 3) = "Service": wsExcel.Columns(3).ColumnWidth = 17
kMax = 3
For kSrv = 1 To 100
    If Mid$(arrROPDOSISRV_Code(kSrv), 1, 1) <> "?" Then
        kMax = kMax + 1
        wsExcel.Cells(Nb, kMax) = arrROPDOSISRV_Code(kSrv)
        arrDetail_Col(kSrv) = kMax

    End If
Next kSrv

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0, " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
  & " where BIATABID = 'ROPDOSGUSR'" _
  & " and SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN' and SSIDOMUIDX = BIATABK1 order by BIATABK1"


Set rsSab = cnsab.Execute(X)


Do While Not rsSab.EOF
    Nb = Nb + 1
    wsExcel.Cells(Nb, 1) = rsSab("BIATABK1")
    X = rsSab("BIATABTXT")
    wsExcel.Cells(Nb, 2) = rsSab("SSIDOMUNIT") 'Mid$(X, 26, 3)
    kSrv = Val(rsSab("SSIDOMUNIT"))
    wsExcel.Cells(Nb, 3) = Trim(arrROPDOSISRV_Code(kSrv))
    For K2 = 1 To 76
        X1 = Trim(Mid$(X, 28 + K2, 1))
        If X1 <> "" Then
            wsExcel.Cells(Nb, arrDetail_Col(K2)) = X1
        End If
    Next K2
    rsSab.MoveNext
Loop

        
For K = 1 To kMax
    If K > 3 Then wsExcel.Columns(K).ColumnWidth = 6.5
    wsExcel.Cells(1, K).Interior.Color = RGB(255, 200, 100)
Next K
For K = 1 To Nb
    wsExcel.Rows(K).RowHeight = 15
    wsExcel.Cells(K, 1).Interior.Color = RGB(255, 200, 100)
    wsExcel.Cells(K, 2).Interior.Color = RGB(255, 255, 128)
    wsExcel.Cells(K, 3).Interior.Color = RGB(255, 255, 128)
Next K
 
'____________________________________________________________________________________

'wbExcel.Worksheets.Add
Set wsExcel = wbExcel.Sheets(3) '   .ActiveSheet
With wsExcel
    .Cells.Font.Name = "Arial"
    .Name = "RO_Qualification"
    .PageSetup.CenterHorizontally = True
    .PageSetup.CenterHeader = "&U&B" & "&""Arial""" & "RO : Qualification"
    .PageSetup.LeftFooter = "&F" & " " & "&A"
    .PageSetup.RightFooter = "&D" & " " & "&T"
    .PageSetup.Orientation = xlLandscape
    .PageSetup.Zoom = False
    .PageSetup.FitToPagesTall = 1
    .PageSetup.FitToPagesWide = 1
    .PageSetup.PrintGridlines = True
End With
With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(0, 0, 255)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .WrapText = True
End With

Nb = 1
wsExcel.Rows(1).RowHeight = 30

wsExcel.Cells(Nb, 1) = "Bâle II": wsExcel.Columns(1).ColumnWidth = 7: wsExcel.Cells(Nb, 1).Interior.Color = RGB(255, 200, 100)
wsExcel.Cells(Nb, 2) = "Qualification": wsExcel.Columns(2).ColumnWidth = 12: wsExcel.Cells(Nb, 2).Interior.Color = RGB(255, 200, 100)
wsExcel.Cells(Nb, 3) = "Libellé": wsExcel.Columns(3).ColumnWidth = 140: wsExcel.Cells(Nb, 3).Interior.Color = RGB(255, 200, 100)

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'ROPDOSQUALB2' order by BIATABK1"
Set rsSab = cnsab.Execute(X)


Do While Not rsSab.EOF
    Nb = Nb + 1
    wsExcel.Cells(Nb, 1) = Val(rsSab("BIATABK1"))
    wsExcel.Cells(Nb, 3) = Trim(rsSab("BIATABTXT"))
    wsExcel.Cells(Nb, 1).Interior.Color = RGB(255, 255, 128)
    wsExcel.Cells(Nb, 2).Interior.Color = RGB(255, 255, 128)
    wsExcel.Cells(Nb, 3).Interior.Color = RGB(255, 255, 128)
    rsSab.MoveNext
Loop

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'ROPDOSQUAL' order by BIATABK1"
Set rsSab = cnsab.Execute(X)


Do While Not rsSab.EOF
    Nb = Nb + 1
    X = Trim(rsSab("BIATABTXT"))
    wsExcel.Cells(Nb, 1) = Val(Mid$(X, 1, 1))
    wsExcel.Cells(Nb, 2) = Trim(rsSab("BIATABK1"))
    wsExcel.Cells(Nb, 3) = Mid$(X, 3, Len(X) - 2)
    rsSab.MoveNext
Loop

Set rsSab = Nothing

wbExcel.SaveAs wFile2

wbExcel.Close

'____________________________________________________________________________________
appExcel.Quit

Set rsSab = Nothing
Set rsSabX = Nothing

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
Call lstErr_AddItem(lstErr, cmdContext, "Export terminé : " & Nb & " enregistrements"): DoEvents

'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub


Private Sub arrYROPDOS0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrYROPDOS0(501)
arrYROPDOS0_Max = 500: arrYROPDOS0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYROPDOS0_GetBuffer(rsSab, xYROPDOS0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmDROPI.fgselect_Display"
        '' Exit Sub
     Else
         arrYROPDOS0_Nb = arrYROPDOS0_Nb + 1
         If arrYROPDOS0_Nb > arrYROPDOS0_Max Then
             arrYROPDOS0_Max = arrYROPDOS0_Max + 50
             ReDim Preserve arrYROPDOS0(arrYROPDOS0_Max)
         End If
         
         arrYROPDOS0(arrYROPDOS0_Nb) = xYROPDOS0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub arrYROPINF0_SQL(lROPDOSID As Long)
Dim V
Dim X As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrYROPINF0(501)
arrYROPINF0_Max = 500: arrYROPINF0_Nb = 0
rsYROPINF0_Init arrYROPINF0(0) ': Processus_Index = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YROPINF0" _
     & " where ROPINFID = " & lROPDOSID _
     & " order by ROPINFIDP,ROPINFIDT,ROPINFIDT2"
     
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYROPINF0_GetBuffer(rsSab, xYROPINF0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmROPINF.fgselect_Display"
        '' Exit Sub
     Else
         arrYROPINF0_Nb = arrYROPINF0_Nb + 1
         If arrYROPINF0_Nb > arrYROPINF0_Max Then
             arrYROPINF0_Max = arrYROPINF0_Max + 50
             ReDim Preserve arrYROPINF0(arrYROPINF0_Max)
         End If
         
         arrYROPINF0(arrYROPINF0_Nb) = xYROPINF0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub lstSelect_Load_1()
Dim I As Long, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_1"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = True
fraSelect_Options_1.Enabled = True
chkSelect_ROPDOSUAMJ.Enabled = True
chkSelect_ROPDOSUAMJ = "0"
chkSelect_ROPDOSGECH.Enabled = True
chkSelect_ROPDOSGECH = "0"
txtSelect_ROPDOSSTA.Enabled = True
txtSelect_ROPDOSGUSR.Enabled = False
chkSelect_ROPDOSGUSR = "1"
chkSelect_ROPDOSGUSR.Enabled = False
fraSelect_Options_1.Visible = True
txtSelect_ROPDOSGUSR.Text = ""

txtSelect_ROPDOSID.Enabled = True
txtSelect_ROPDOSXID.Enabled = True
txtSelect_ROPINFGTXT.Enabled = True
txtSelect_ROPDOSXDOM.Enabled = True
txtSelect_ROPDOSXAPP.Enabled = True
txtSelect_ROPDOSGPRV.Enabled = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub

Private Sub lstSelect_Load_2()
Dim I As Long, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_2"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = False
cmdSelect_Ok_Click
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub


Private Sub lstSelect_Load_6()
Dim I As Long, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_6"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = False
'cmdSelect_Ok_Click
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub

Private Sub lstSelect_Load_7()
Dim I As Long, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_7"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True

fraSelect_Options_1.Visible = True
fraSelect_Options_1.Enabled = True
chkSelect_ROPDOSGECH.Enabled = False
chkSelect_ROPDOSGECH = "1"
Call DTPicker_Set(txtSelect_ROPDOSGECH_Max, DSys) '_SuivantO)

txtSelect_ROPDOSSTA.Enabled = False
txtSelect_ROPDOSGUSR.Enabled = False
chkSelect_ROPDOSGUSR = "1"
chkSelect_ROPDOSGUSR.Enabled = False
fraSelect_Options_1.Visible = True

txtSelect_ROPDOSID.Enabled = False
txtSelect_ROPDOSXID.Enabled = False
txtSelect_ROPINFGTXT.Enabled = False
txtSelect_ROPDOSXDOM.Enabled = True
txtSelect_ROPDOSXAPP.Enabled = True
txtSelect_ROPDOSGPRV.Enabled = False
txtSelect_ROPDOSGUSR.Enabled = DROPI_Aut.Xspécial
If DROPI_Aut.Xspécial Then
    txtSelect_ROPDOSGUSR = ""
Else
    txtSelect_ROPDOSGUSR = usrName_UCase
End If
'cmdSelect_Ok_Click
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
blnControl = True
If cboSelect_SQL.ListIndex = 0 And mcboSelect_SQL_ListIndex > 0 Then cboSelect_SQL.ListIndex = mcboSelect_SQL_ListIndex: Exit Sub

If fraExport.Visible Then
    fraExport.Visible = False
    Exit Sub
End If
cmdUpdate_Dossier.ListIndex = -1
cmdUpdate.ListIndex = -1

Select Case SSTab1.Tab
    Case Is = 0
    
        If fraDossier.Visible Then
        
            If fraUpdate_PJ.Visible Then
                tabDossier.Tab = 3
                tabDossier.Caption = ""
                fraUpdate_PJ.Visible = False
                fraUpdate_B.Visible = False
                tabDossier.Tab = 1
                cmdDossier_Ok.Visible = False
                fraDossier_cmd.Visible = True
                Exit Sub
            End If
            If fraUpdate_B.Visible Then
                'If cmdDossier_Ok.Visible Then
                '    tabDossier.Tab = 2
                'Else
                '    fraUpdate_B.Visible = False: cmdUpdate.Visible = False
                '    tabDossier.Tab = 1
                'End If
                    tabDossier.Tab = 2
                    tabDossier.Caption = ""
                    cmdUpdate_12.Visible = False
                    cmdUpdate_22.Visible = False
                    cmdUpdate_32.Visible = False

                    cmdDossier_Ok.Visible = False
                    cmdDossier_Ok_Close.Visible = False
                    fraDossier_cmd.Visible = True
                    fraUpdate_B.Visible = False: cmdUpdate.Visible = False
                    tabDossier.Tab = 1
                Exit Sub
            End If
            If lstUpdate_ROPINFMAIL.Visible Then
                tabDossier.Tab = 3
                tabDossier.Caption = ""
                lstUpdate_ROPINFMAIL.Visible = False: lstUpdate_ROPINFMAIL_Display.Visible = False: libUpdate_ROPINFMAIL.Visible = False
                lstUpdate_ROPINFMAIL_CC.Visible = False: lstUpdate_ROPINFMAIL_CC_Display.Visible = False: libUpdate_ROPINFMAIL_CC.Visible = False
                tabDossier.Tab = 1
                cmdDossier_Ok.Visible = False
                fraDossier_cmd.Visible = True
                Exit Sub
            End If
            If lstUpdate_ROPDOSQUAL.Visible Then
                 lstUpdate_ROPDOSQUAL.Visible = False
                 tabDossier.Tab = 1
                 cmdDossier_Ok.Visible = False
                 fraDossier_cmd.Visible = True
                 Exit Sub
             End If
            'If fraDossier.Visible Then
             '    fraDossier.Visible = False
                 'tabDossier.Tab = 1
             '    Exit Sub
             'End If
             If cmdDossier_Ok.Visible Then
                 cmdDossier_Ok.Visible = False
                 fraDossier_cmd.Visible = True
                 If tabDossier.Tab = 0 Then
                    cmdUpdate_Dossier.Visible = True
                 Else
                    tabDossier.Tab = 1
                 End If
                 Exit Sub
             End If
           fraDossier.Visible = False
        Else
            Unload Me
        End If
    Case Is = 1
        If fraParam_Update.Visible Then
            fraParam_Update.Visible = False
        Else
            SSTab1.Tab = 0
        End If
     Case Is = 2
        If fraAut_Update.Visible Then
            fraAut_Update.Visible = False
        Else
            SSTab1.Tab = 0
        End If
   Case Else
        Unload Me
End Select
End Sub





Private Sub cboSelect_SQL_Click()
cmdUpdate_Dossier.Visible = False
cmdSelect_SQL_K = Trim(Mid$(cboSelect_SQL, 1, 2))
If blnControl Then
    cmdSelect_Reset
    Me.Enabled = False: Me.MousePointer = vbHourglass
    'fraDossier_B.Visible = False
    'libDossier_ROPINFGTXT.Visible = False
    'lstUpdate_ROPINFMAIL.Visible = False
    'fraSelect_Options_1.Visible = False
    'fraDossier.Visible = False
    Select Case cmdSelect_SQL_K
        Case "0": cmdSelect_Ok_Click
        Case "1": lstSelect_Load_1
        Case "1M": cmdSelect_Ok_Click
        Case "2": lstSelect_Load_2
        Case "2M": lstSelect_Load_2
        Case "6": lstSelect_Load_6
        Case "7", "7@": lstSelect_Load_7
        Case "S1": cmdSelect_Ok_Click
    End Select
    Me.Enabled = True: Me.MousePointer = 0
End If

End Sub


Private Sub chkSelect_ROPDOSGECH_Click()
If chkSelect_ROPDOSGECH = "1" Then
    txtSelect_ROPDOSGECH_Max.Visible = True
Else
    txtSelect_ROPDOSGECH_Max.Visible = False
End If


End Sub

Private Sub chkSelect_ROPDOSGECH_GotFocus()
cmdSelect_Reset

End Sub


Private Sub chkSelect_ROPDOSGUSR_Click()
If chkSelect_ROPDOSGUSR = "1" Then
    txtSelect_ROPDOSGUSR.Visible = True
Else
    txtSelect_ROPDOSGUSR.Visible = False
End If

End Sub


Private Sub chkSelect_ROPDOSGUSR_GotFocus()
cmdSelect_Reset

End Sub


Private Sub chkSelect_ROPDOSUAMJ_Click()
If chkSelect_ROPDOSUAMJ = "1" Then
    txtSelect_ROPDOSUAMJ.Visible = True
Else
    txtSelect_ROPDOSUAMJ.Visible = False
End If

End Sub

Private Sub chkSelect_Update_B_Click()
End Sub

Private Sub chkSelect_ROPDOSUAMJ_GotFocus()
cmdSelect_Reset

End Sub

Private Sub chkUpdate_ROPINFGPRV_GotFocus()
'cmdSelect_Reset

End Sub


Private Sub chkUpdate_ROPINFMAIL_Click()
If chkUpdate_ROPINFMAIL = "1" Then
    fraUpdate_ROPINFMAIL.Visible = True
Else
    fraUpdate_ROPINFMAIL.Visible = False
End If
End Sub

Private Sub cmdAut_Update_Ok_Click()
Dim V
Dim blnOk As Boolean
Dim K As Integer, kSrv As Integer

Me.Enabled = False: Me.MousePointer = vbHourglass
'___________________________________________________________
App_Debug = "cmdAut_Update_Ok"
blnOk = False

If IsNumeric(Mid$(txtAut_ROPDOSGUSR_SRV, 2, 2)) Then
    kSrv = Val(Mid$(txtAut_ROPDOSGUSR_SRV, 2, 2))
    If arrROPDOSISRV_K1(kSrv) <> "" Then
        blnOk = True
        '''''Mid$(newAut.BIATABTXT, 26, 3) = arrROPDOSISRV_K1(kSrv)
    End If
End If
'K = InStr(28, newAut.BIATABTXT, "R")
'If K > 0 Then
'    blnOk = True
'Else
'    K = InStr(28, newAut.BIATABTXT, "D")
'    If K > 0 Then
'        blnOk = True
'    Else
'        K = InStr(28, newAut.BIATABTXT, "C")
'        If K > 0 Then
'            blnOk = True
'        Else
'            K = InStr(28, newAut.BIATABTXT, "X")
'            If K > 0 Then blnOk = True
'        End If
'    End If
'End If

If Not blnOk Then
    Call lstErr_Clear(lstErr, cmdContext, "? préciser le service de ce collaborateur"): DoEvents
Else
    Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement " & App_Debug): DoEvents
    'Mid$(newAut.BIATABTXT, 26, 3) = "S" & Format$(kSrv, "00")
    cmdAut_Update_Ok_Transaction
    
    Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement" & App_Debug): DoEvents
End If
'___________________________________________________________
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdAut_Update_Quit_Click()
fraAut_Update.Visible = False

End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdDossier_Ok_04()
Dim V
blnUpdate_Sucess = False
App_Debug = "cmdDossier_Ok_04"
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement " & App_Debug): DoEvents

If IsNull(fraDossier_Control) Then
    oldYROPDOS0 = newYROPDOS0
    If IsNull(fraDétail_Update_Control) Then
    newYROPINF0.ROPINFGNAT = "P"
    'Création Dossier : insert YROPDOS0 et YROPINF0 (description)
    '---------------------------------------------------------------

        V = cmdDossier_Ok_Transaction("Insert")
        If Not IsNull(V) Then
            MsgBox V, vbCritical, Me.Name & " : cmddétail_Update_Ok" & App_Debug
            Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
        Else
            If blnSendMail Then Call cmdSendMail("D")
            blnControl = False: cboSelect_SQL.ListIndex = mcboSelect_SQL_ListIndex: blnControl = True
            cmdSelect_SQL_K = 1: cmdSelect_SQL_X1 = 1
            Call cmdSelect_SQL_1(newYROPDOS0.ROPDOSID)
            'Call cmdUpdate_Reset
            blnUpdate_Sucess = True
        End If
    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement" & App_Debug): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Function cmdDossier_Ok_Transaction(lFct As String)
Dim V, X As String, xSQL As String
Dim xSet As String, xWhere As String
Dim Nb As Long, K As Integer
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdDossier_Ok_Transaction"
'-------------------------------------------------------
cmdDossier_Ok_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

'________________________________________________________________________________
Call cmdDossier_Ok_GSRV(newYROPDOS0.ROPDOSIUSR, newYROPDOS0.ROPDOSISRV)
Call cmdDossier_Ok_GSRV(newYROPDOS0.ROPDOSGUSR, newYROPDOS0.ROPDOSGSRV)
mailYROPDOS0 = newYROPDOS0: mailYROPINF0 = newYROPINF0

Select Case lFct
    Case "Update": V = sqlYROPDOS0_Update(newYROPDOS0, oldYROPDOS0, True)
    Case "Insert":
                If blnDossierModèle Or blnDossierReprise Then
                    newYROPDOS0.ROPDOSID = DossierModèle_ROPDOSID
                Else
                    V = sqlROPDOSID_Init("ROPDOSID", newYROPDOS0.ROPDOSID)
                    currentROPDOSID = newYROPDOS0.ROPDOSID
                End If
                If IsNull(V) Then
                    Call cmdDossier_Ok_GSRV(newYROPINF0.ROPINFGUSR, newYROPINF0.ROPINFGSRV)
                    V = sqlYROPDOS0_Insert(newYROPDOS0)    ' dossier & description
                    If IsNull(V) Then
                        newYROPINF0.ROPINFID = newYROPDOS0.ROPDOSID
                        newYROPINF0.ROPINFUVER = 1
                        V = sqlYROPINF0_Insert(newYROPINF0)
                        mailYROPDOS0 = newYROPDOS0: mailYROPINF0 = newYROPINF0
'_________________________________________________________________________________

                        For K = 2 To dupYROPINF0_Nb
                            newYROPINF0 = dupYROPINF0(K)
                            If newYROPINF0.ROPINFGECH > mailYROPINF0.ROPINFGECH Then
                                newYROPINF0.ROPINFGECH = mailYROPINF0.ROPINFGECH
                            End If
                            If Trim(newYROPINF0.ROPINFGUSR) = "?" Then
                                Select Case newYROPINF0.ROPINFGNAT
                                    Case "F": newYROPINF0.ROPINFGUSR = newYROPDOS0.ROPDOSIUSR
                                    Case Else: newYROPINF0.ROPINFGUSR = newYROPDOS0.ROPDOSGUSR
                                End Select
                            End If
                            newYROPINF0.ROPINFID = newYROPDOS0.ROPDOSID
                            Call cmdDossier_Ok_GSRV(newYROPINF0.ROPINFGUSR, newYROPINF0.ROPINFGSRV)
                            V = sqlYROPINF0_Insert(newYROPINF0)
                        Next K
'_________________________________________________________________________________
                        
                    End If
                End If
    Case "Delete": V = sqlYROPDOS0_Delete(oldYROPDOS0): mailYROPDOS0 = oldYROPDOS0
    Case Else: V = "? fct non traitée : " & lFct
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case cmdUpdate_K
    Case 14
        newYROPINF0 = arrYROPINF0(1)
        newYROPINF0.ROPINFGECH = newYROPDOS0.ROPDOSGECH
        V = sqlYROPINF0_Update(newYROPINF0, arrYROPINF0(1), True)
        If Not IsNull(V) Then GoTo Error_MsgBox
        arrYROPINF0(1) = newYROPINF0

    Case 24, 34, 54:
        xSet = " set ROPINFSTAD = '" & mROPINFSTAD_Set & "'" & " , ROPINFSTAK = '" & mROPINFSTAK_Set & "'"
        
        xWhere = " where ROPINFID = " & oldYROPDOS0.ROPDOSID _
       & " and ROPINFSTAD = '" & mROPINFSTAD_Where & "'"

        V = sqlYROPINF0_Requête("update ", xSet, xWhere)
        If Not IsNull(V) Then GoTo Error_MsgBox
    Case 44:
        
        xWhere = " where ROPINFID = " & oldYROPDOS0.ROPDOSID
       
        V = sqlYROPINF0_Requête("delete from ", "", xWhere)
        If Not IsNull(V) Then GoTo Error_MsgBox
End Select
'________________________________________________________________________________

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
        If blnIncidentSignificatif_Mail Then Call cmdSendMail("IncidentSignificatif")
        If blnSécurité_Mail Then Call cmdSendMail("Sécurité")
        fgSelect_YROPDOS0_Read
    End If
    
    cmdDossier_Ok_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function
Public Function cmdDossier_Ok_Transaction_Duplication()
Dim V, X As String, xSQL As String
Dim xSet As String, xWhere As String
Dim Nb As Long, K As Integer
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdDossier_Ok_Transaction_Duplication"
'-------------------------------------------------------
cmdDossier_Ok_Transaction_Duplication = Null
'________________________________________________________________________________
mailYROPDOS0 = oldYROPDOS0: mailYROPINF0 = arrYROPINF0(1)

arrYROPINF0_SQL (Val(Mid$(lstUpdate_Modèle.Text, 1, 4)))
'________________________________________________________________________________

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox


'_________________________________________________________________________________

For K = 2 To arrYROPINF0_Nb
    newYROPINF0 = arrYROPINF0(K)
    newYROPINF0.ROPINFGECH = mailYROPINF0.ROPINFGECH

    If Trim(newYROPINF0.ROPINFGUSR) = "?" Then
        Select Case newYROPINF0.ROPINFGNAT
            Case "F": newYROPINF0.ROPINFGUSR = oldYROPDOS0.ROPDOSIUSR
            Case Else: newYROPINF0.ROPINFGUSR = oldYROPDOS0.ROPDOSGUSR
        End Select
    End If
    newYROPINF0.ROPINFID = oldYROPDOS0.ROPDOSID
    V = sqlYROPINF0_Insert(newYROPINF0)
Next K
'_________________________________________________________________________________
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________

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
        If blnSendMail Then Call cmdSendMail("D")
        cmdSelect_SQL_K = 1: cmdSelect_SQL_X1 = 1
        Call cmdSelect_SQL_1(newYROPDOS0.ROPDOSID)

    End If
    
    cmdDossier_Ok_Transaction_Duplication = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function

Public Function cmdParam_Update_Ok_Transaction()
Dim V, X As String, xSQL As String
Dim xSet As String, xWhere As String
Dim Nb As Long, wseq As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdParam_Update_Ok_Transaction"
'-------------------------------------------------------
cmdParam_Update_Ok_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case cmdParam_SQL_K
    Case "Update": V = sqlYBIATAB0_Update(newParam, oldParam)
            If blnROPDOSMAIL Then
                '$JPL 2013-10-01newROPDOSMAIL = oldROPDOSMAIL
                '$JPL 2013-10-01newROPDOSMAIL.BIATABTXT = Trim(txtROPDOSMAIL)
                '$JPL 2013-10-01V = sqlYBIATAB0_Update(newROPDOSMAIL, oldROPDOSMAIL)
            End If
    Case "Insert":
                If Trim(oldParam.BIATABID) = "ROPINFGTXT" Then
                    V = sqlROPDOSID_Init("ROPINFGTXT_$", wseq)
                    newParam.BIATABK2 = Format$(wseq, "000000000000")
                    Else
                    If IsNull(sqlYBIATAB0_Read(Trim(newParam.BIATABID), Trim(newParam.BIATABK1), Trim(newParam.BIATABK2), oldParam.BIATABTXT)) Then
                        V = "Existe déjà : " & Trim(newParam.BIATABID) & " " & Trim(newParam.BIATABK1) & " " & Trim(newParam.BIATABK2)
                    Else
                        V = Null
                    End If
                End If
                If IsNull(V) Then
                    V = sqlYBIATAB0_Insert(newParam)    ' dossier & description
                End If
    Case "Delete": V = sqlYBIATAB0_Delete(oldParam)
    Case Else: V = "? fct non traitée : " & cmdParam_SQL_K
End Select
If Not IsNull(V) Then GoTo Error_MsgBox

Select Case Trim(oldParam.BIATABID)
    Case "ROPDOSXDOM": lstParam_ROPDOSXDOM_Load mlstParam_ListIndex: sqlYBIATAB0_cboID "ROPDOSXDOM", txtUpdate_ROPDOSXDOM

    Case "ROPDOSXAPP": lstParam_ROPDOSXAPP_Load (oldParam.BIATABK1), mlstParam_ListIndex
    Case "ROPINFGTXT": lstParam_ROPINFGTXT_Load mlstParam_ListIndex
    Case "ROPDOSQUAL": lstParam_ROPDOSQUAL_Load mlstParam_ListIndex
    Case "ROPDOSQUALB2": lstParam_ROPDOSQUALB2_Load mlstParam_ListIndex
    Case "ROPDOSISRV": Form_Init_RODOSISRV: Form_Init_RODOSGUSR
End Select

'________________________________________________________________________________

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
       fraParam_Update.Visible = False
    End If
    
    cmdParam_Update_Ok_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function

Public Function cmdAut_Update_Ok_Transaction()
Dim V, X As String, xSQL As String
Dim xSet As String, xWhere As String
Dim Nb As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdAut_Update_Ok_Transaction"
'-------------------------------------------------------
cmdAut_Update_Ok_Transaction = Null

Mid$(newAut.BIATABTXT, 1, 7) = "*******"
If chkAut_ROPDOSGUSR = "1" Then Mid$(newAut.BIATABTXT, 1, 1) = "D"
If chkAut_ROPINFGUSR = "1" Then Mid$(newAut.BIATABTXT, 2, 1) = "A"
If chkAut_ROPDOSGUSR_P = "1" Then Mid$(newAut.BIATABTXT, 3, 1) = "P"
If chkAut_ROPDOSGUSR_H = "1" Then Mid$(newAut.BIATABTXT, 4, 1) = "H"
If chkAut_ROPDOSGUSR_Q = "1" Then Mid$(newAut.BIATABTXT, 5, 1) = "Q"
If chkAut_ROPDOSGUSR_E = "1" Then Mid$(newAut.BIATABTXT, 6, 1) = "E"
If chkAut_ROPDOSGUSR_I = "1" Then Mid$(newAut.BIATABTXT, 7, 1) = "I"

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case cmdAut_SQL_K
    Case "Update":
                    V = sqlYBIATAB0_Update(newAut, oldAut)
    Case "Insert":
                    V = sqlYBIATAB0_Insert(newAut)
    Case "Delete": V = sqlYBIATAB0_Delete(oldAut)
    Case Else: V = "? fct non traitée : " & cmdAut_SQL_K
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________

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
       fraAut_Update.Visible = False
    End If
    
    cmdAut_Update_Ok_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function

Public Function cmdDétail_Update_Ok_Transaction(lFct As String)
Dim V, X As String, xSQL As String
Dim xSet As String, xWhere As String
Dim Archive_Folder As String, Archive_File As String
Dim App_Event As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdDétail_Update_Ok_Transaction"
App_Event = "Begintrans"
'-------------------------------------------------------
cmdDétail_Update_Ok_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Call FEU_ROUGE
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
mailYROPDOS0 = oldYROPDOS0: mailYROPINF0 = newYROPINF0: savYROPINF0 = newYROPINF0

App_Event = lFct

If cmdUpdate_K = "05" Then
    newYROPINF0.ROPINFGUSR = newFileExtension
Else
    Call cmdDossier_Ok_GSRV(newYROPINF0.ROPINFGUSR, newYROPINF0.ROPINFGSRV)
End If
Select Case lFct
    Case "Update":
            V = sqlYROPINF0_Update(newYROPINF0, oldYROPINF0, True)
    Case "Insert":
            If blnROPINFIDT_Insérer Then
                cmdDétail_Update_Ok_Transaction_Insérer
            Else
                V = sqlYROPINF0_Insert(newYROPINF0)
            End If
            
    Case "Delete": V = sqlYROPINF0_Delete(oldYROPINF0): mailYROPINF0 = oldYROPINF0
    Case Else: V = "? fct non traitée : " & lFct
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case cmdUpdate_K
    Case "05"
        newFileName = newDirPath & "\" & newYROPINF0.ROPINFID & "_" & newYROPINF0.ROPINFIDP _
            & "_" & newYROPINF0.ROPINFIDT & "_" & newYROPINF0.ROPINFIDT2 & "." & newFileExtension
            
        App_Event = "MkDir " & newDirPath

        If Not msFileSystem.FolderExists(newDirPath) Then MkDir newDirPath
        App_Event = "CopyFile " & oldFileName & vbCrLf & newFileName
        If Dir(newFileName) <> "" Then Kill newFileName

        msFileSystem.CopyFile oldFileName, newFileName
        
        If Mid$(oldFileName, 1, 8) = "C:\Temp\" Then
            Archive_Folder = "C:\Temp\Archive"
            App_Event = "MkDir " & Archive_Folder
           If Not msFileSystem.FolderExists(Archive_Folder) Then MkDir Archive_Folder
            X = Mid$(oldFileName, 8, Len(oldFileName) - 7)
            X = Replace(X, "\", "_")
            Archive_File = Archive_Folder & "\" & newYROPINF0.ROPINFID & "_" & DSys & "_" & time_Hms & "_" & X
            App_Event = "MoveFile " & oldFileName & vbCrLf & Archive_File
           msFileSystem.MoveFile oldFileName, Archive_File
        End If

    ''Case "11", "12", "13"
      ''  newYROPDOS0 = oldYROPDOS0
        '$JPL_20071004 newYROPDOS0.ROPDOSIUSR = newYROPINF0.ROPINFGUSR
      ''  cmdDossier_Ok_04_ISRV
      ''  V = sqlYROPDOS0_Update(newYROPDOS0, oldYROPDOS0, True)
      ''  mailYROPDOS0 = newYROPDOS0
      ''  If Not IsNull(V) Then GoTo Error_MsgBox
   Case "22", "32", "52":
        xSet = " set ROPINFSTA = '" & mROPINFSTA_Set & "'" & " , ROPINFSTAK = '" & mROPINFSTAK_Set & "'"
        
        xWhere = " where ROPINFID = " & oldYROPINF0.ROPINFID _
       & " and ROPINFIDP  = " & oldYROPINF0.ROPINFIDP _
       & " and ROPINFIDT  = " & oldYROPINF0.ROPINFIDT _
       & " and ROPINFSTA = '" & mROPINFSTA_Where & "'"

        V = sqlYROPINF0_Requête("update ", xSet, xWhere)
        If Not IsNull(V) Then GoTo Error_MsgBox
    Case "42":
        
        xWhere = " where ROPINFID = " & oldYROPINF0.ROPINFID _
       & " and ROPINFIDP  = " & oldYROPINF0.ROPINFIDP _
       & " and ROPINFIDT  = " & oldYROPINF0.ROPINFIDT
       
        V = sqlYROPINF0_Requête("delete from ", "", xWhere)
        If Not IsNull(V) Then GoTo Error_MsgBox
    Case "23", "33", "53":
        xSet = " set ROPINFSTA = '" & mROPINFSTA_Set & "'" & " , ROPINFSTAK = '" & mROPINFSTAK_Set & "'"
        
        xWhere = " where ROPINFID = " & oldYROPINF0.ROPINFID _
       & " and ROPINFIDP  = " & oldYROPINF0.ROPINFIDP _
       & " and ROPINFSTA = '" & mROPINFSTA_Where & "'"

        V = sqlYROPINF0_Requête("update ", xSet, xWhere)
        If Not IsNull(V) Then GoTo Error_MsgBox
    Case "43":
        
        xWhere = " where ROPINFID = " & oldYROPINF0.ROPINFID _
       & " and ROPINFIDP  = " & oldYROPINF0.ROPINFIDP
       
        V = sqlYROPINF0_Requête("delete from ", "", xWhere)
        If Not IsNull(V) Then GoTo Error_MsgBox
End Select
'________________________________________________________________________________


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V & vbCrLf & App_Event, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        'cmdUpdate_Reset
        fgSelect_YROPDOS0_Read
    End If
    
    cmdDétail_Update_Ok_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Call FEU_VERT
End Function
Public Function cmdDétail_Update_Ok_Transaction_Insérer()
Dim V, K As Integer
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdDétail_Update_Ok_Transaction_Insérer"
'-------------------------------------------------------
cmdDétail_Update_Ok_Transaction_Insérer = Null
newYROPINF0.ROPINFIDT = oldYROPINF0.ROPINFIDT
'________________________________________________________________________________
V = sqlYROPINF0_Delete_GE(oldYROPINF0)
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
V = sqlYROPINF0_Insert(newYROPINF0)
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
For K = 1 To arrYROPINF0_Nb
    If arrYROPINF0(K).ROPINFID = newYROPINF0.ROPINFID _
    And arrYROPINF0(K).ROPINFIDP = newYROPINF0.ROPINFIDP _
    And arrYROPINF0(K).ROPINFIDT >= newYROPINF0.ROPINFIDT Then

        arrYROPINF0(K).ROPINFIDT = arrYROPINF0(K).ROPINFIDT + 1
        V = sqlYROPINF0_Insert(arrYROPINF0(K))
        If Not IsNull(V) Then GoTo Error_MsgBox
    End If
Next K

Exit Function

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    cmdDétail_Update_Ok_Transaction_Insérer = V

End Function

Public Function cmdSelect_SQL_6_Transaction()
Dim V, K As Long

On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdSelect_SQL_6_Transaction"
'-------------------------------------------------------
cmdSelect_SQL_6_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

If oldYROPDOS0.ROPDOSSTAK <> xYROPDOS0.ROPDOSSTAK Then
    newYROPDOS0 = xYROPDOS0
    V = sqlYROPDOS0_Update(newYROPDOS0, oldYROPDOS0, False)
End If


If Not IsNull(V) Then GoTo Error_MsgBox

'________________________________________________________________________________
    For K = 1 To arrYROPINF0_Nb
        If selYROPINF0(K).ROPINFSTAK <> arrYROPINF0(K).ROPINFSTAK Then
            V = sqlYROPINF0_Update(selYROPINF0(K), arrYROPINF0(K), False)
        End If
    Next K
'________________________________________________________________________________

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
    
    cmdSelect_SQL_6_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function

Private Sub cmdDétail_Update_Ok()
Dim V, X As String
App_Debug = "cmdDossier_Ok"

If cmdUpdate_Fct = "Delete" Then
    V = Null
Else
    V = fraDétail_Update_Control
End If

If IsNull(V) Then

    V = cmdDétail_Update_Ok_Transaction(cmdUpdate_Fct)
    If Not IsNull(V) Then
        MsgBox V, vbCritical, Me.Name & " : cmddétail_Update_Ok" & App_Debug
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    Else
        If blnSendMail Then Call cmdSendMail(" ")
        xYROPDOS0 = oldYROPDOS0
        fraDossier_Display
    End If
End If
End Sub
Private Sub cmdDétail_Update_Ok_Insert()
Dim V, X As String, K As Integer
App_Debug = "cmdDétail_Update_Ok_Insert"

If IsNull(fraDétail_Update_Control) Then

    Select Case cmdUpdate_K
        Case "01", "05"
                        For K = 1 To arrYROPINF0_Nb
                            If newYROPINF0.ROPINFID = arrYROPINF0(K).ROPINFID _
                            And newYROPINF0.ROPINFIDP = arrYROPINF0(K).ROPINFIDP _
                            And newYROPINF0.ROPINFIDT = arrYROPINF0(K).ROPINFIDT Then
                                If newYROPINF0.ROPINFIDT2 < arrYROPINF0(K).ROPINFIDT2 Then newYROPINF0.ROPINFIDT2 = arrYROPINF0(K).ROPINFIDT2
                            End If
                        Next K
                        newYROPINF0.ROPINFIDT2 = newYROPINF0.ROPINFIDT2 + 1
         Case "02"
                        For K = 1 To arrYROPINF0_Nb
                            If newYROPINF0.ROPINFID = arrYROPINF0(K).ROPINFID _
                            And newYROPINF0.ROPINFIDP = arrYROPINF0(K).ROPINFIDP Then
                                If newYROPINF0.ROPINFIDT < arrYROPINF0(K).ROPINFIDT Then newYROPINF0.ROPINFIDT = arrYROPINF0(K).ROPINFIDT
                            End If
                        Next K
                        newYROPINF0.ROPINFIDT = newYROPINF0.ROPINFIDT + 1
                        newYROPINF0.ROPINFIDT2 = 1
         Case "03"
                        For K = 1 To arrYROPINF0_Nb
                            If newYROPINF0.ROPINFID = arrYROPINF0(K).ROPINFID Then
                                If newYROPINF0.ROPINFIDP < arrYROPINF0(K).ROPINFIDP Then newYROPINF0.ROPINFIDP = arrYROPINF0(K).ROPINFIDP
                            End If
                        Next K
                        newYROPINF0.ROPINFIDP = newYROPINF0.ROPINFIDP + 1
                        newYROPINF0.ROPINFIDT = 0
                        newYROPINF0.ROPINFIDT2 = 1
   End Select
   
   newYROPINF0.ROPINFSTA = " "
    V = cmdDétail_Update_Ok_Transaction("Insert")
    If Not IsNull(V) Then
        MsgBox V, vbCritical, Me.Name & " : cmddétail_Update_Ok" & App_Debug
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    Else
        If blnSendMail Then Call cmdSendMail(" ")
        xYROPDOS0 = oldYROPDOS0
        fraDossier_Display

    End If
End If

End Sub

Private Sub cmdDossier_Ok_14()
Dim V, X As String
App_Debug = "cmdDossier_Ok"

Select Case cmdUpdate_K
    Case "14Q": V = Null
    Case Else: V = fraDossier_Control
End Select
    If IsNull(V) Then
    
        V = cmdDossier_Ok_Transaction("Update")
        If Not IsNull(V) Then
            MsgBox V, vbCritical, Me.Name & " : cmdDossier_Ok" & App_Debug
            Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
        Else
            oldYROPDOS0 = newYROPDOS0
            If blnSendMail Then Call cmdSendMail(" ")
            xYROPDOS0 = oldYROPDOS0
            fraDossier_Display
            lstUpdate_ROPDOSQUAL.Visible = False
        End If
    End If

End Sub


Private Sub cmdDossier_Mail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
    lstUpdate_ROPINFMAIL.TopIndex = 0
    lstUpdate_ROPINFMAIL_CC.TopIndex = 0
    lstUpdate_ROPINFMAIL.Visible = True: lstUpdate_ROPINFMAIL_Display.Visible = True: libUpdate_ROPINFMAIL.Visible = True
    lstUpdate_ROPINFMAIL_CC.Visible = True: lstUpdate_ROPINFMAIL_CC_Display.Visible = True: libUpdate_ROPINFMAIL_CC.Visible = True
    cmdDossier_Ok.Visible = True: fraDossier_cmd.Visible = False
    fraUpdate_ROPINFMAIL.Visible = False
    tabDossier.Tab = 3
    tabDossier.Caption = "Envoi mail"
    cmdUpdate_K = 98
Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdDossier_Print_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

prtDROPI_Open 1, "Risques Opérationnels"
cmdPrint0_Dossier
prtDROPI_Close 1

Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdExport_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Export en cours ......"): DoEvents

YROPDOS0_Export

fraExport.Visible = False

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdExport_Quit_Click()
    fraExport.Visible = False
End Sub

Private Sub cmdParam_Update_Ok_Click()
Dim V
App_Debug = "cmdParam_Update_Ok"
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement " & App_Debug): DoEvents

If IsNull(fraParam_Update_Control) Then cmdParam_Update_Ok_Transaction

Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement" & App_Debug): DoEvents

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Update_Quit_Click()
fraParam_Update.Visible = False

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
Dim I As Integer

blnControl = False
usrColor_Set

cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
blncmdOk_Visible = False: blncmdSave_Visible = False
currentAction = ""
SSTab1.Tab = 0

blnAuto = False
blnAuto_Ok = False
fraSelect_Options_1.Visible = False
fraDossier.Visible = False
cmdSelect_Ok.Caption = "Extraire les mouvements"

libRéférenceInterne = ""

'$JPL 20120502
Select Case cboSelect_SQL.ListCount
    Case Is = 0: mcboSelect_SQL_ListIndex = -1
    Case Is = 1: mcboSelect_SQL_ListIndex = 0
    Case Is > 1: mcboSelect_SQL_ListIndex = 1
End Select
cboSelect_SQL.ListIndex = mcboSelect_SQL_ListIndex
lstSelect_Load_1

'If cboSelect_SQL.ListCount > 0 Then
'    If currentROPDOSISRV_Rôle = "R" Or currentROPDOSISRV_Rôle = "D" Or currentROPDOSISRV_Rôle = "H" Then
'        cboSelect_SQL.ListIndex = 1: lstSelect_Load_1
'    Else
'        cboSelect_SQL.ListIndex = 1: lstSelect_Load_1
'        'cboSelect_SQL.ListIndex = 0
'    End If
'End If

False_Aut.Avis = False
False_Aut.Comptabiliser = False
False_Aut.Consulter = False
False_Aut.Rapprocher = False
False_Aut.Saisir = False
False_Aut.Swift = False
False_Aut.Valider = False
False_Aut.Virement = False
False_Aut.Xspécial = DROPI_Aut.Xspécial


libDossier_ROPINFID.ForeColor = &H606060
libDossier_ROPDOSUUSR.ForeColor = &H606060

dirListBox.PATH = "C:\Temp"
chkUpdate_ROPINFMAIL.Value = "0"
fraUpdate_ROPINFMAIL.Visible = False
blnControl = True
End Sub
Public Sub Form_Init()
Dim V, X As String
Dim xZBASTAB0 As typeZBASTAB0
Dim K As Integer
Dim rsMDB As New ADODB.Recordset

Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents
fgSelect_FormatString = fgSelect.FormatString
fgSelect.Visible = False


SSTab1.BackColor = &HC0F0FF
blnControl = False
lstW.Visible = False

fraDossier.Visible = False
Set fraDossier.Container = fraSelect.Container
fraDossier.Left = fraSelect.Width - fraDossier.Width - 400
fraDossier.Top = fraSelect.Top + 230
'cmdUpdate.Top = 600
'cmdUpdate.Height = 1968
'cmdUpdate.Width = 3564
'cmdUpdate.Left = 5760
'cmdUpdate_Dossier.Top = 600
'cmdUpdate_Dossier.Height = 1968
'cmdUpdate_Dossier.Width = 3564
'cmdUpdate_Dossier.Left = 5760



vDsys = dateImp(DSys)
fraSelect.Enabled = DROPI_Aut.Consulter
cmdSelect_Ok.Visible = False
cmdDossier_Ok_Close.Visible = False

txtSelect_ROPDOSGECH_Max.Visible = False
chkSelect_ROPDOSGECH.Value = "0"
Call DTPicker_Set(txtSelect_ROPDOSGECH_Max, DateAdd_AMJ("d", 7, DSys))
mROPDOSIAMJ_Min = DateAdd_AMJ("d", -7, DSys)

txtSelect_ROPDOSUAMJ.Visible = False
chkSelect_ROPDOSUAMJ.Value = "0"
Call DTPicker_Set(txtSelect_ROPDOSUAMJ, DSys)

chkSelect_ROPDOSGUSR = "0"
txtSelect_ROPDOSGUSR.Visible = False
txtSelect_ROPDOSGUSR = usrName_UCase

Form_Init_RODOSISRV
Form_Init_RODOSGUSR

sqlYBIATAB0_cboID "ROPDOSXDOM", txtSelect_ROPDOSXDOM
txtSelect_ROPDOSXDOM.AddItem "             - tous les domaines"
txtSelect_ROPDOSXDOM.ListIndex = 0
txtSelect_ROPDOSGPRV.Clear
txtSelect_ROPDOSGPRV.AddItem "  - tous"
'txtSelect_ROPDOSGPRV.AddItem "U - " & usrName_UCase
txtSelect_ROPDOSGPRV.AddItem "V - " & currentROPDOSISRV_Nom
txtSelect_ROPDOSGPRV.AddItem "W - public"
txtSelect_ROPDOSGPRV.ListIndex = 0

cboSelect_SQL.Clear
If DROPI_Aut.Consulter Then
    cboSelect_SQL.AddItem "1  - Extraire"   '" & usrName_UCase & "' + '" & currentROPDOSISRV_Nom & "'"
    'cboSelect_SQL.AddItem "1X - Rechercher un texte"
End If
If DROPI_Aut.Saisir Then
    cboSelect_SQL.AddItem "0  - Saisie fiche RO"
    '$JPL à revoir cboSelect_SQL.AddItem "2  - Ouverture d'un dossier"
    '$JPL à revoir  cboSelect_SQL.AddItem "7  - Echéancier (impression)"
    '$JPL à revoir  cboSelect_SQL.AddItem "7@ - Echéancier (mail)"
End If
If DROPI_Aut.Xspécial Then
    '$JPL à revoir  cboSelect_SQL.AddItem "1M - Liste des modèles"
    '$JPL à revoir  cboSelect_SQL.AddItem "2M - Création d'un modèle"
    'cboSelect_SQL.AddItem "2R - Reprise d'un dossier"
    cboSelect_SQL.AddItem "6  - Màj statut des dossiers"
    'cboSelect_SQL.AddItem "9$ - spécial JPL : restucturation fichiers "
    cboSelect_SQL.AddItem "S1 - statistiques  "
    ''''cboSelect_SQL.AddItem "JPL - spécial JPL  "
End If

'_____________________________________________________________________________
sqlYBIATAB0_cboID "ROPDOSSTA", txtSelect_ROPDOSSTA
txtSelect_ROPDOSSTA.AddItem "* - tous"
txtSelect_ROPDOSSTA.ListIndex = 0
sqlYBIATAB0_cboID "ROPDOSSTA", txtUpdate_ROPDOSSTA
sqlYBIATAB0_cboID "ROPDOSXDOM", txtUpdate_ROPDOSXDOM
txtUpdate_ROPDOSXDOM.AddItem "?"
sqlYBIATAB0_cboID "ROPDOSGNAT", txtSelect_ROPDOSGNAT
txtSelect_ROPDOSGNAT.AddItem "* - tous"
txtSelect_ROPDOSGNAT.ListIndex = 0
sqlYBIATAB0_cboID "ROPDOSGNAT", txtUpdate_ROPDOSGNAT
txtUpdate_ROPDOSGNAT.AddItem "?"
sqlYBIATAB0_cboID "ROPDOSGPRV", txtUpdate_ROPDOSGPRV
txtUpdate_ROPDOSGPRV.AddItem "?"
sqlYBIATAB0_cboID "ROPDOSGGRA", txtUpdate_ROPDOSGGRA
sqlYBIATAB0_cboID "ROPDOSGPRI", txtUpdate_ROPDOSGPRI
sqlYBIATAB0_cboID "ROPDOSQUAL", txtUpdate_ROPDOSQUAL

''cboZMNURUT0_Load_Prod txtUpdate_ROPINFGUSR
'sqlYBIATAB0_cboID "ROPINFMAIL", txtUpdate_ROPINFMAIL
sqlYBIATAB0_cboID "ROPINFSTA", txtUpdate_ROPINFSTA
sqlYBIATAB0_cboID "ROPINFGNAT", txtUpdate_ROPINFGNAT

txtUpdate_ROPINFSTA.Locked = True
txtUpdate_ROPDOSSTA.Locked = True
'_____________________________________________________________________________

paramROPDOS_Path = paramServer("\\ROPDOS\") & paramEnvironnement & "\"
paramROPDOS_Path_DROPI = paramServer("\\ROPDOS_DROPI\" & paramEnvironnement & "\")

'_____________________________________________________________________________
lstParam_ROPINFGTXT_Load 0
lstParam_ROPDOSXDOM_Load 0
rsYROPINF0_Init zYROPINF0
cmdReset
txtUpdate_ROPDOSGTXT.ForeColor = vbBlue ' &H4000&
'_____________________________________________________________________________
cmdUpdate.Visible = False
'cmdUpdate.BackColor = &HF6FFF6
'cmdUpdate.ForeColor = &H4000&     'vbBlue

cmdUpdate_Dossier.Visible = False
'cmdUpdate_Dossier.BackColor = &HF6FFF6
'cmdUpdate_Dossier.ForeColor = &H4000&     'vbBlue
libDossier_ROPINFGTXT.Visible = False
libDossier_ROPINFGTXT.ForeColor = vbBlue

lstUpdate_ROPINFMAIL.Visible = False
lstUpdate_ROPINFMAIL.ForeColor = vbBlue
lstUpdate_ROPINFMAIL_Display.Visible = False: libUpdate_ROPINFMAIL.Visible = False
lstUpdate_ROPINFMAIL_Display.Top = lstUpdate_ROPINFMAIL.Top
lstUpdate_ROPINFMAIL_Display.Height = lstUpdate_ROPINFMAIL.Height
lstUpdate_ROPINFMAIL_Display.Left = lstUpdate_ROPINFMAIL.Left + lstUpdate_ROPINFMAIL.Width + 100
lstUpdate_ROPINFMAIL_Display.BackColor = &HD0FFD0
lstUpdate_ROPINFMAIL.BackColor = &HD0FFD0
'libUpdate_ROPINFMAIL.Left = lstUpdate_ROPINFMAIL_Display.Left


lstUpdate_ROPINFMAIL_CC.Visible = False
lstUpdate_ROPINFMAIL_CC.ForeColor = vbBlue
lstUpdate_ROPINFMAIL_CC_Display.Visible = False: libUpdate_ROPINFMAIL_CC.Visible = False
lstUpdate_ROPINFMAIL_CC_Display.Top = lstUpdate_ROPINFMAIL_CC.Top
lstUpdate_ROPINFMAIL_CC_Display.Height = lstUpdate_ROPINFMAIL_CC.Height
lstUpdate_ROPINFMAIL_CC_Display.Left = lstUpdate_ROPINFMAIL_CC.Left + lstUpdate_ROPINFMAIL_CC.Width + 100
lstUpdate_ROPINFMAIL_CC_Display.BackColor = &HD0FFFF
lstUpdate_ROPINFMAIL_CC.BackColor = &HD0FFFF
'libUpdate_ROPINFMAIL_CC.Left = lstUpdate_ROPINFMAIL_CC_Display.Left


txtUpdate_ROPINFGTXT_0.ForeColor = vbMagenta
txtUpdate_ROPINFGTXT_0.BackColor = &HC0FFC0
lblUpdate_ROPINFIDTL.ForeColor = vbRed
fraUpdate_B.ForeColor = vbMagenta
libDossier_ROPDOSID.ForeColor = vbRed
libDossier_ROPDOSIUSR.ForeColor = vbMagenta
libDossier_ROPDOSGUSR.ForeColor = vbBlue
txtUpdate_ROPINFIDTL.BackColor = vbRed
txtUpdate_ROPINFIDTL.ForeColor = vbWhite
lstUpdate_Modèle.BackColor = &H80FF80
chkUpdate_ROPINFMAIL_D.ForeColor = vbBlue
chkUpdate_ROPINFMAIL_P.ForeColor = vbBlue
chkUpdate_ROPINFMAIL_A.ForeColor = vbBlue
chkUpdate_ROPINFMAIL_I.ForeColor = vbBlue
chkUpdate_ROPINFMAIL_U.ForeColor = vbBlue
SSTab1.Visible = True
'lstUpdate_Modèle_Load

If fraParam.Visible Then fraParam_Reset
If fraAut.Visible Then fraAut_Reset
'____________________________________________________________________________________________
fraUpdate_PJ.Top = 480
fraUpdate_PJ.Left = 240

txtUpdate_ROPDOSQUAL.Visible = blnROPDOSQUAL
lblUpdate_ROPDOSQUAL.Visible = blnROPDOSQUAL
lstUpdate_ROPDOSQUAL.Visible = False
lstUpdate_ROPDOSQUAL.Top = 50
lstUpdate_ROPDOSQUAL.Left = 50
lstUpdate_ROPDOSQUAL.Width = 4600
lstUpdate_ROPDOSQUAL.Height = fraDossier.Height - 100
'Set lstUpdate_ROPDOSQUAL.Container = fraDossier 'libDossier_ROPINFGTXT.Container

For I = 0 To txtUpdate_ROPDOSQUAL.ListCount - 1
    txtUpdate_ROPDOSQUAL.ListIndex = I
    lstUpdate_ROPDOSQUAL.AddItem txtUpdate_ROPDOSQUAL.Text
Next I
lstParam_ROPDOSQUAL_Load 0
lstParam_ROPDOSQUALB2_Load 0

'________________________________________________________________________________________
fraParam.BorderStyle = 0
SSTab2.ForeColor = vbMagenta
lblParam_ROPDOSXDOM.ForeColor = &HC00000
lblParam_ROPDOSXAPP.ForeColor = &HC00000
lblParam_ROPINFGTXT.ForeColor = &HC00000
lblParam_ROPDOSQUAL.ForeColor = &HC00000
lblParam_ROPDOSQUALB2.ForeColor = &HC00000
'________________________________________________________________________________________

Set fraExport.Container = frmDROPI
fraExport.Top = 600
fraExport.Left = 120

Select Case Val(Mid$(DSys, 5, 2))
    Case Is < 4
        Call DTPicker_Set(txtExport_AMJMIN, (Mid$(DSys, 1, 4) - 1) & "1001")
        Call DTPicker_Set(txtExport_AMJMAX, (Mid$(DSys, 1, 4) - 1) & "1231")
    Case Is < 7
        Call DTPicker_Set(txtExport_AMJMIN, Mid$(DSys, 1, 4) & "0101")
        Call DTPicker_Set(txtExport_AMJMAX, Mid$(DSys, 1, 4) & "0331")
    Case Is < 10
        Call DTPicker_Set(txtExport_AMJMIN, Mid$(DSys, 1, 4) & "0401")
        Call DTPicker_Set(txtExport_AMJMAX, Mid$(DSys, 1, 4) & "0630")
    Case Else
        Call DTPicker_Set(txtExport_AMJMIN, Mid$(DSys, 1, 4) & "0701")
        Call DTPicker_Set(txtExport_AMJMAX, Mid$(DSys, 1, 4) & "0930")
End Select
'________________________________________________________________________________________

fgDetail.Rows = 1
'fgDetail.ColWidth(0) = 100
fgDetail.ColWidth(0) = 9000
txtDetail.Width = fgDetail.ColWidth(0)

txtSelect_txt.Width = fgSelect.ColWidth(2)
'If cboSelect_SQL.ListIndex = 0 Then cmdSelect_Ok_Click
txtROPDOSMAIL.ForeColor = vbRed
End Sub

Private Sub cmdDossier_Ok_Close_Click()
Dim V
App_Debug = "cmdDossier_Ok_Close"
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement " & App_Debug): DoEvents

blnSendMail = False
blnUpdate_Sucess = False
Select Case cmdUpdate_K
    Case "04": cmdDossier_Ok_04:
               If blnUpdate_Sucess Then
                    fraDossier_B.Visible = True
                    cmdUpdate.Clear
                    cmdUpdate.AddItem "05 - Ajouter une pièce jointe"
                    cmdUpdate.ListIndex = 0
               End If
    Case "02":
        cmdDétail_Update_Ok_Insert
        oldYROPINF0 = savYROPINF0
        xYROPINF0 = savYROPINF0
        fraDétail_Display
        mROPINFSTA_Value = "F": mROPINFSTA_Set = "+": mROPINFSTA_Where = " ": mROPINFSTAK_Set = "V"
        blnSendMail = True
        If oldYROPINF0.ROPINFGNAT <> "F" Then
            cmdUpdate_K = "22"
            cmdDétail_Update_Ok
        Else
            cmdUpdate_K = "23"
            cmdUpdate_Init_23_Ok
        End If
    Case "12":
        cmdDétail_Update_Ok
        oldYROPINF0 = newYROPINF0
        mROPINFSTA_Value = "F": mROPINFSTA_Set = "+": mROPINFSTA_Where = " ": mROPINFSTAK_Set = "V"
        blnSendMail = True
        If oldYROPINF0.ROPINFGNAT <> "F" Then
            cmdUpdate_K = "22"
            cmdDétail_Update_Ok
        Else
            cmdUpdate_K = "23"
            cmdUpdate_Init_23_Ok
        End If
    Case "13":
        cmdDétail_Update_Ok
        oldYROPINF0 = newYROPINF0
        blnSendMail = True
        cmdUpdate_K = "23"
        cmdUpdate_Init_23_Ok
End Select

libDossier_ROPINFGTXT.Visible = False
lstUpdate_ROPINFMAIL.Visible = False: lstUpdate_ROPINFMAIL_Display.Visible = False: libUpdate_ROPINFMAIL.Visible = False
lstUpdate_ROPINFMAIL_CC.Visible = False: lstUpdate_ROPINFMAIL_CC_Display.Visible = False: libUpdate_ROPINFMAIL_CC.Visible = False

Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement" & App_Debug): DoEvents

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_New_Click()
If cboSelect_SQL.ListIndex = 0 Then cboSelect_SQL.ListIndex = -1
cboSelect_SQL.ListIndex = 0

End Sub

Private Sub cmdUpdate_01_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "Ajouter une NOTE"): DoEvents
cmdUpdate_Init_01
Me.Enabled = True: Me.MousePointer = 0
txtUpdate_ROPINFGTXT.SetFocus

End Sub

Private Sub cmdUpdate_02_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "Ajouter une ACTION"): DoEvents
Call cmdUpdate_Init_02("A")
Me.Enabled = True: Me.MousePointer = 0
txtUpdate_ROPINFGTXT.SetFocus

End Sub


Private Sub cmdUpdate_05_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "Ajouter une NOTE"): DoEvents
cmdUpdate_Init_05
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdUpdate_12_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdUpdate_K = Mid$(cmdUpdate_12.Caption, 1, 2)
cmdUpdate_Exe
Me.Enabled = True: Me.MousePointer = 0
txtUpdate_ROPINFGTXT.SetFocus

End Sub

Private Sub cmdUpdate_22_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdUpdate_K = Mid$(cmdUpdate_22.Caption, 1, 2)
cmdUpdate_Exe
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdUpdate_32_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdUpdate_K = Mid$(cmdUpdate_32.Caption, 1, 2)
cmdUpdate_Exe
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdUpdate_Dossier_Click()
Dim blnReset As Boolean
If cmdUpdate_Dossier.ListIndex < 0 Then Exit Sub

cmdUpdate_K = Trim(Mid$(cmdUpdate_Dossier, 1, 3))
'cmdUpdate.Clear
'cmdUpdate.AddItem cmdUpdate_Dossier
blnReset = False
'cmdUpdate_Dossier.Visible = False
Select Case cmdUpdate_K
    Case "03": cmdUpdate_Init_03
    Case "14": cmdUpdate_Init_14
    Case "14Q": cmdUpdate_Init_14Q
    Case "24": cmdUpdate_Init_24
    Case "34": cmdUpdate_Init_34
    Case "44": cmdUpdate_Init_44
    Case "54": cmdUpdate_Init_54
    Case "64": cmdUpdate_Init_64
End Select

If blnReset Then fraDossier.Visible = False
End Sub


Private Sub dirListBox_Change()
filDoc.PATH = dirListBox.PATH
filDoc.Pattern = "*.*"
End Sub


Public Sub lstParam_ROPDOSXAPP_Load(lK1 As String, lIndex As Long)
Dim X As String, X12 As String
On Error Resume Next
lstParam_ROPDOSXAPP.Clear
X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID =  'ROPDOSXAPP' " _
    & " and BIATABK1 = '" & lK1 & "' order by BIATABK2"
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    X12 = rsSab("BIATABK2")
    lstParam_ROPDOSXAPP.AddItem X12 & " - " & Trim(Mid$(rsSab("BIATABTXT"), 1, 24))
    rsSab.MoveNext
Loop

If lstParam_ROPDOSXAPP.ListCount > 0 Then lstParam_ROPDOSXAPP.ListIndex = lIndex
End Sub

Public Sub lstParam_ROPDOSGUSR_Load(lK1 As String)
Dim xK1 As String, xUsr As String
Dim xHab As String
On Error Resume Next
'_________________________________________________________________
kParam_ROPDOSISRV = Val(Mid$(lK1, 2, 2))
lstParam_ROPDOSGUSR.Clear
X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0, " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
  & " where BIATABID = 'ROPDOSGUSR'" _
  & " and SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN' and SSIDOMUIDX = BIATABK1 and SSIDOMPRFK <> 'X' order by BIATABK1"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    xHab = ""
    xK1 = UCase$(rsSab("BIATABK1"))
    xUsr = Trim(xK1)
    X = rsSab("BIATABTXT")
    If rsSab("SSIDOMUNIT") = lK1 Then xHab = "*"
    If Mid$(X, 28 + kParam_ROPDOSISRV, 1) <> " " Then xHab = Mid$(X, 28 + kParam_ROPDOSISRV, 1)
    If xHab <> "" Then lstParam_ROPDOSGUSR.AddItem xHab & " - " & xUsr
    rsSab.MoveNext
Loop
        
'_____________________________________________________________________
End Sub

Public Sub lstParam_ROPDOSXDOM_Load(lIndex As Long)
Dim X As String, X12 As String
On Error Resume Next
lstParam_ROPDOSXDOM.Clear
X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID =  'ROPDOSXDOM' order by BIATABK1"
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    X12 = rsSab("BIATABK1")
    lstParam_ROPDOSXDOM.AddItem X12 & " - " & Trim(Mid$(rsSab("BIATABTXT"), 1, 24))
    rsSab.MoveNext
Loop

If lstParam_ROPDOSXDOM.ListCount > 0 Then lstParam_ROPDOSXDOM.ListIndex = lIndex
End Sub

Public Sub lstParam_ROPDOSQUAL_Load(lIndex As Long)
Dim X As String, X3 As String
On Error Resume Next

X = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID =  'ROPDOSQUAL'"
Set rsSab = cnsab.Execute(X)
arrROPDOSQUAL_Nb = 0
ReDim arrROPDOSQUAL(rsSab("Tally") + 1)
ReDim arrROPDOSQUAL_Code(rsSab("Tally") + 1)
lstParam_ROPDOSQUAL.Clear
X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID =  'ROPDOSQUAL' order by BIATABK1"
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    X3 = Mid$(rsSab("BIATABK1"), 1, 3)
    arrROPDOSQUAL_Nb = arrROPDOSQUAL_Nb + 1
    arrROPDOSQUAL_Code(arrROPDOSQUAL_Nb) = X3
    arrROPDOSQUAL(arrROPDOSQUAL_Nb) = Trim(rsSab("BIATABTXT"))
    
    lstParam_ROPDOSQUAL.AddItem X3 & " - " & arrROPDOSQUAL(arrROPDOSQUAL_Nb)
    rsSab.MoveNext
Loop

If lstParam_ROPDOSQUAL.ListCount > 0 Then lstParam_ROPDOSQUAL.ListIndex = lIndex '0
End Sub

Public Sub lstParam_ROPDOSQUALB2_Load(lIndex As Long)
Dim X As String, X1 As String, K As Integer
On Error Resume Next
lstParam_ROPDOSQUALB2.Clear
X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID =  'ROPDOSQUALB2' order by BIATABK1"
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    X1 = Mid$(rsSab("BIATABK1"), 1, 1)
    K = Val(X1)
    arrROPDOSQUALB2(K) = X1 & " - " & Trim(rsSab("BIATABTXT"))
    lstParam_ROPDOSQUALB2.AddItem arrROPDOSQUALB2(K)
    rsSab.MoveNext
Loop

If lstParam_ROPDOSQUALB2.ListCount > 0 Then lstParam_ROPDOSQUALB2.ListIndex = lIndex
End Sub


Public Sub lstParam_ROPINFGTXT_Load(lIndex As Long)
Dim X As String, X12 As String * 12
Dim mBIATABK1 As String
On Error Resume Next

mBIATABK1 = ""
lstParam_ROPINFGTXT.Clear
libDossier_ROPINFGTXT.Clear

X = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID =  'ROPINFGTXT'"
Set rsSab = cnsab.Execute(X)
ReDim arrROPINFGTXT_BIATABK2(rsSab("Tally"))

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID =  'ROPINFGTXT' order by BIATABK1,BIATABTXT"
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    X12 = rsSab("BIATABK1")
    lstParam_ROPINFGTXT.AddItem X12 & " - " & Trim(rsSab("BIATABTXT"))
    arrROPINFGTXT_BIATABK2(lstParam_ROPINFGTXT.ListCount - 1) = rsSab("BIATABK2")
    If mBIATABK1 <> X12 Then
        mBIATABK1 = X12
        libDossier_ROPINFGTXT.AddItem "___________ " & UCase(Trim(X12)) & " ___________________________________"
    End If
    libDossier_ROPINFGTXT.AddItem Trim(rsSab("BIATABTXT"))
    rsSab.MoveNext
Loop

If lstParam_ROPINFGTXT.ListCount > 0 Then lstParam_ROPINFGTXT.ListIndex = lIndex
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


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
mnuExport.Visible = blnExportation_xlsx
mnuExport_Param.Visible = blnExportation_xlsx
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Dim Msg As String
Dim I As Integer

Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case SSTab1.Tab
    Case 0: Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
End Select
Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

Me.Enabled = False: Me.MousePointer = vbHourglass
    fraDossier.Visible = False

blnOk = True
Call lstErr_Clear(lstErr, cmdContext, "> SAB_CDR_cmdSelect_Ok ........"): DoEvents
cmdSelect_Ok.Visible = False
cmdUpdate_Dossier.Visible = False



blnDossierModèle = False
blnDossierReprise = False
DoEvents
If blnOk Then
    cmdSelect_SQL_X1 = Mid$(cboSelect_SQL, 1, 1)
    cmdSelect_SQL_K = Trim(Mid$(cboSelect_SQL, 1, 2))
    Select Case cmdSelect_SQL_K
        Case "0": cmdSelect_SQL_0
        Case "1": Call cmdSelect_SQL_1(0)
        Case "1M":    blnDossierModèle = True: Call cmdSelect_SQL_1(0)
        Case "2": Call cmdSelect_SQL_1(0)
        Case "2M": cmdSelect_SQL_2M
        Case "6": cmdSelect_SQL_6
        Case "7", "7@": cmdSelect_SQL_7: blnOk = False
        Case "9$": cmdSelect_SQL_JPL
        Case "S1": cmdSelect_SQL_S1
        'Case "JP": cmdSelect_SQL_JPL

    End Select

End If
If Not blnOk Then

'Else
    cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
    cmdSelect_Ok.BackColor = &HC0F0FF    '&HE0FFFF
    fraSelect_Options_1.BackColor = &HC0F0FF   ' &HE0FFFF  ' &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options_1, fraSelect_Options_1.BackColor)
    fraSelect_Options_1.Enabled = True

End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_CDR_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
cmdSelect_Ok.Visible = True

End Sub

Public Sub cmdSelect_Reset()
If blnControl Then
    lstErr.Clear
    fgSelect.Visible = False
    cmdSelect_Ok.Visible = True
    
    fraDossier.Visible = False
    fraDossier_B.Visible = False
    'txtUpdate_ROPDOSGTXT.BackColor = &HF0FFFF
    fgDetail.Visible = False
    
    fraUpdate_B.Visible = False
    cmdUpdate.Visible = False
    'txtUpdate_ROPINFGTXT.BackColor = &HF0FFFF
    
    libDossier_ROPINFGTXT.Visible = False
    lstUpdate_Modèle.Visible = False
    lstUpdate_ROPINFMAIL.Visible = False: lstUpdate_ROPINFMAIL_Display.Visible = False:    libUpdate_ROPINFMAIL.Visible = False
    lstUpdate_ROPINFMAIL_CC.Visible = False: lstUpdate_ROPINFMAIL_CC_Display.Visible = False:    libUpdate_ROPINFMAIL_CC.Visible = False
    fraUpdate_PJ.Visible = False
    tabDossier.Tab = 3: tabDossier.Caption = ""
    tabDossier.Tab = 2: tabDossier.Caption = ""
    tabDossier.Tab = 1: tabDossier.Caption = ""

End If

End Sub

Private Sub cmdSelect_SQL_1(lROPDOSID As Long)
Dim V, X As String
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String, blnAnd As Boolean, filtreROPDOSGUSR As String, filtreROPINFGUSR As String
Dim blnROPDOSGUSR As Boolean, wId As Long
Dim blnFiltre As Boolean
On Error GoTo Error_Handler

fgSelect.Clear
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_1"): DoEvents

currentAction = "cmdSelect_SQL_1"
xWhere = "": xAnd = "": filtreROPDOSGUSR = ""
blnROPDOSGUSR = False
blnFiltre = False
If lROPDOSID > 0 Then
    currentROPDOSID = lROPDOSID
    xWhere = " where ROPDOSId =" & lROPDOSID
Else
    Select Case cmdSelect_SQL_K
        Case "2", "1M":
                Select Case currentROPDOSISRV
                    Case "_S31": xAnd = " and   ROPDOSID < 1000"
                    Case "_S40": xAnd = " and   ROPDOSID < 1000"
                    Case Else: xAnd = " and   ROPDOSID < 1000 and ROPDOSID > 9"
                End Select
       Case "2M":
            currentROPDOSID = Trim(txtSelect_ROPDOSID)
            If currentROPDOSID <> "" Then xWhere = " and   ROPDOSId =" & currentROPDOSID
        Case Else: blnFiltre = True
    End Select
End If

If blnFiltre Then
    Select Case Mid$(txtSelect_ROPDOSGPRV, 1, 1)
        Case "V": cmdSelect_SQL_K = "1V"
        Case "W": cmdSelect_SQL_K = "1W"
    End Select
    currentROPDOSID = Trim(txtSelect_ROPDOSID)
    If currentROPDOSID <> "" And Val(currentROPDOSID) >= 1000 Then
        xWhere = " Where   ROPDOSId =" & currentROPDOSID
        GoTo cmdSelect_SQL_1_Ok
    Else
        X = Trim(txtSelect_ROPDOSXID)
        If X <> "" Then
           xWhere = " Where   ROPDOSXID= '" & X & "'"
            GoTo cmdSelect_SQL_1_Ok
        Else
            X = Trim(txtSelect_ROPINFGTXT)
            If X <> "" Then Call cmdSelect_SQL_1X: Exit Sub
        End If
    
    End If
    
    xWhere = xWhere & " and   ROPDOSID >= 1000"
    
    X = Mid$(txtSelect_ROPDOSSTA, 1, 1)
    If X <> "*" Then xWhere = xWhere & " and   ROPDOSSTA = '" & X & "'"
    
    X = Mid$(txtSelect_ROPDOSGNAT, 1, 1)
    If X <> "*" Then xWhere = xWhere & " and   ROPDOSGNAT = '" & X & "'"
    
    Call DTPicker_Control(txtSelect_ROPDOSGECH_Max, WAMJMax)
    
    If chkSelect_ROPDOSGECH = "1" Then
        xWhere = xWhere & " and   ROPDOSgech <= '" & WAMJMax & "'"
    End If
    
       
    Call DTPicker_Control(txtSelect_ROPDOSUAMJ, WAMJMax)
    
    If chkSelect_ROPDOSUAMJ = "1" Then
        xWhere = xWhere & " and   ROPDOSCAMJ >= '" & WAMJMax & "'"
    End If
 
    X = Trim(Mid$(txtSelect_ROPDOSXDOM, 1, 12))
    If X <> "" Then
        xAnd = xAnd & " and   ROPDOSXDOM = '" & X & "'"
        X = Trim(Mid$(txtSelect_ROPDOSXAPP, 1, 12))
        If X <> "" Then xAnd = xAnd & " and ROPDOSXAPP = '" & X & "'"
    End If
    
End If
'____________________________________________________________________________________
'usrName_UCase & " + " & currentROPDOSISRV
'____________________________________________________________________________________

Select Case cmdSelect_SQL_K
    Case "1": filtreROPDOSGUSR = ""
               blnROPDOSGUSR = True
    Case "1V": filtreROPDOSGUSR = "  and ( ROPDOSISRV = '" & currentROPDOSISRV & "' or ROPDOSGSRV = '" & currentROPDOSISRV & "')"
               filtreROPINFGUSR = " and ROPINFGSRV = '" & currentROPDOSISRV & "'"
               blnROPDOSGUSR = True
    Case "1W": filtreROPDOSGUSR = " and  ROPDOSGPRV = 'W'"
End Select
'____________________________________________________________________________________

cmdSelect_SQL_1_Ok:
'==================

If xWhere <> "" Then
    Mid$(xWhere, 1, 6) = " where"
Else
    If xAnd <> "" Then Mid$(xAnd, 1, 6) = " where"
End If

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 " & xWhere & xAnd & filtreROPDOSGUSR
Set rsSab = cnsab.Execute(xSQL)
ReDim selYROPDOS0(100): selYROPDOS0_Nb = 0
Do While Not rsSab.EOF
        If selYROPDOS0_Nb >= UBound(selYROPDOS0) Then ReDim Preserve selYROPDOS0(selYROPDOS0_Nb + 100)
        selYROPDOS0_Nb = selYROPDOS0_Nb + 1
        V = rsYROPDOS0_GetBuffer(rsSab, selYROPDOS0(selYROPDOS0_Nb))
    rsSab.MoveNext
Loop
'____________________________________________________________________________________

cmdSelect_SQL_1_YROPDOS0
'____________________________________________________________________________________

If currentROPDOSID <> "" Then fraDossier_Display
Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub cmdSelect_SQL_1X()
Dim V, X As String
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String
Dim wId As Long
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_1X"): DoEvents

cmdSelect_SQL_K = "1X"
currentAction = "cmdSelect_SQL_1X"
X = Trim(txtSelect_ROPINFGTXT)
    If X = "" Then
        Call lstErr_AddItem(lstErr, cmdContext, "? Préciser un texte à rechercher"): DoEvents
        Exit Sub
End If
xWhere = " where ROPINFGTXT like '%" & X & "%'"
xSQL = "select  distinct(ROPINFID) from " & paramIBM_Library_SABSPE & ".YROPINF0 " _
    & xWhere
Set rsSab = cnsab.Execute(xSQL)
X = ""
Do While Not rsSab.EOF
    wId = rsSab("ROPINFID")
    If wId > 0 Then X = X & wId & ","
    rsSab.MoveNext
Loop
ReDim selYROPDOS0(100): selYROPDOS0_Nb = 0

If X <> "" Then
    xWhere = "Where ROPDOSID in(" & Mid$(X, 1, Len(X) - 1) & ")"
   xSQL = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 " & xWhere & xAnd
   Set rsSab = cnsab.Execute(xSQL)
   Do While Not rsSab.EOF
        If selYROPDOS0_Nb >= UBound(selYROPDOS0) Then ReDim Preserve selYROPDOS0(selYROPDOS0_Nb + 100)
        selYROPDOS0_Nb = selYROPDOS0_Nb + 1
        V = rsYROPDOS0_GetBuffer(rsSab, selYROPDOS0(selYROPDOS0_Nb))
       rsSab.MoveNext
   Loop
End If
'____________________________________________________________________________________
cmdSelect_SQL_1_YROPDOS0

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_2()
Dim V, X As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_2"): DoEvents

currentAction = "cmdSelect_SQL_2"

blnSelect_Update_B_Display = True
Me.Enabled = True: Me.MousePointer = 0
''''fraDossier.Left = fraDossier_Left

cmdUpdate_K = "04"
currentROPDOSID = oldYROPDOS0.ROPDOSID
xYROPDOS0 = oldYROPDOS0
fraDossier_Display_YROPDOS0
fraDossier_Display_YROPINF0

'fraDossier_Display
'fraDétail_Display
cmdUpdate_Init_04

'fraUpdate_B.Enabled = True
cmdUpdate.Visible = False
cmdDossier_Ok.Visible = True: fraDossier_cmd.Visible = False
cmdDossier_Ok_Close.Caption = "Enregistrer + ajouter une pièce jointe"
cmdDossier_Ok_Close.Visible = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdUpdate_Click()
Dim xItem As String, K As Integer
If cmdUpdate.ListIndex < 0 Then Exit Sub
'xItem = Mid$(cmdUpdate, 1, Len(cmdUpdate))
cmdUpdate_K = Mid$(cmdUpdate, 1, 2)
cmdUpdate_Exe
End Sub


Private Sub cmdDossier_Ok_Click()
Dim V, X As String
App_Debug = "cmdDossier_Ok"
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement " & App_Debug): DoEvents

Select Case cmdUpdate_K
    Case "01", "02", "03": cmdDétail_Update_Ok_Insert
    Case "04": cmdDossier_Ok_04
    Case "05":
        If Len(rtfPJ.Text) > 0 Then
            oldFileName = "C:\temp\DROPI.rtf"
            If Dir(oldFileName) <> "" Then Kill oldFileName
            newDirPath = paramROPDOS_Path & oldYROPINF0.ROPINFID
            X = InputBox("Préciser le nom du document", "DROPI : Pièce jointe")
            If X = "" Then X = DSYS_Time
            newFileName = X & ".rtf"
            newFileExtension = "rtf"
            txtUpdate_ROPINFGTXT = newFileName
            fraUpdate_PJ.Visible = False
            txtUpdate_ROPINFGTXT.Locked = False
            rtfPJ.SaveFile oldFileName
        End If
        If Trim(newFileName) <> "" Then cmdDétail_Update_Ok_Insert
    Case "11", "12", "13": cmdDétail_Update_Ok
    Case "14", "14Q": cmdDossier_Ok_14
    Case "74": cmdDossier_Ok_Transaction_Duplication
    Case "98": mailYROPDOS0 = oldYROPDOS0: mailYROPINF0 = arrYROPINF0(1)
                Call cmdSendMail("X")
                cmdUpdate_Reset
End Select


Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement" & App_Debug): DoEvents
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdDossier_Quit_Click()

cmdContext_Quit


End Sub


Private Sub cmdUpdate_Reset()
fraUpdate_PJ.Visible = False
fraDossier_B.Visible = False
fraDossier_cmd.Visible = False
cmdDossier_Ok.Visible = False
cmdDossier_Ok_Close.Visible = False
blnSelect_Update_B_Display = False
fraDossier.Visible = False
lstUpdate_ROPDOSQUAL.Visible = False
End Sub

Private Sub DriveListBox_Change()
On Error Resume Next
dirListBox.PATH = DriveListBox.Drive ' .PATH
End Sub

Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
If fgDetail.Rows > 1 Then
     
    'Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
    
    fgDetail.Col = 1
    K = Val(fgDetail.Text)
    xYROPINF0 = arrYROPINF0(K)
    oldYROPINF0 = xYROPINF0
    currentYROPINF0 = xYROPINF0
    cmdUpdate_Init_K = xYROPINF0.ROPINFGNAT
'___________________________________________________________________________________
    Processus_Index = 0: Action_Index = 0: Action_Suivante_Index = 0
    
    For K = 1 To arrYROPINF0_Nb
        If xYROPINF0.ROPINFIDP = arrYROPINF0(K).ROPINFIDP Then
            If arrYROPINF0(K).ROPINFGNAT = "P" Then Processus_Index = K
            If arrYROPINF0(K).ROPINFGNAT = "A" Or arrYROPINF0(K).ROPINFGNAT = "F" Then
                If xYROPINF0.ROPINFIDT = arrYROPINF0(K).ROPINFIDT Then
                    Action_Index = K
                Else
                    If xYROPINF0.ROPINFIDT < arrYROPINF0(K).ROPINFIDT And Action_Suivante_Index = 0 Then
                        mailYROPINF0_Suivant = arrYROPINF0(K)
                        Action_Suivante_Index = K: Exit For
                    End If
               End If
                
           End If
        End If
    Next K
'___________________________________________________________________________________

    If xYROPINF0.ROPINFGNAT = "J" And Button = 1 Then
        Call fraDétail_Display_PJ_FileName(oldYROPINF0.ROPINFGUSR, True)
    Else
        cmdUpdate_Init
        fraDétail_Display
        fraUpdate_B.Visible = True
        tabDossier.Tab = 2
        Select Case oldYROPINF0.ROPINFGNAT
            Case "P":   tabDossier.Caption = "Processus " & oldYROPINF0.ROPINFIDP
            Case "A":   tabDossier.Caption = oldYROPINF0.ROPINFIDP & "-Action " & oldYROPINF0.ROPINFIDT
            Case "N":   tabDossier.Caption = oldYROPINF0.ROPINFIDP & "-Note " & oldYROPINF0.ROPINFIDT
            Case "J":   tabDossier.Caption = oldYROPINF0.ROPINFIDP & "-Pièce Jointe " & oldYROPINF0.ROPINFIDT
            Case Else:  tabDossier.Caption = oldYROPINF0.ROPINFGNAT & oldYROPINF0.ROPINFIDT
        End Select
    End If
End If

End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Integer, xSQL As String

If fgSelect.Rows > 1 Then
    fgSelect.Visible = False
    
    
    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    
    fgSelect.Col = 0
    
    K = InStr(fgSelect.Text, "-")
    currentROPDOSID = Val(Mid$(fgSelect.Text, 1, K - 1))
    fgSelect_YROPDOS0_Read
    
    fgSelect.Visible = True
    fgSelect.LeftCol = 0
End If

End Sub

Private Sub filDoc_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
oldFileName = filDoc.PATH & "\" & filDoc.FileName
newDirPath = paramROPDOS_Path & oldYROPINF0.ROPINFID
newFileName = filDoc.FileName
newFileExtension = fileName_Extension(filDoc.FileName)
'txtUpdate_ROPINFGUSR = newFileExtension
txtUpdate_ROPINFGTXT = filDoc.FileName
fraUpdate_PJ.Visible = False
txtUpdate_ROPINFGTXT.Locked = False

cmdDossier_Ok_Click
Me.Enabled = True: Me.MousePointer = 0
On Error Resume Next
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
    Case Is = 27: KeyCode = 0: cmdContext_Quit
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





Private Sub Label1_Click()

End Sub

Private Sub libDossier_ROPINFGTXT_Click()
On Error Resume Next
txtUpdate_ROPINFGTXT = txtUpdate_ROPINFGTXT & libDossier_ROPINFGTXT
txtUpdate_ROPINFGTXT.SetFocus
txtUpdate_ROPINFGTXT.SelStart = Len(txtUpdate_ROPINFGTXT)
libDossier_ROPINFGTXT.Visible = False
End Sub



Private Sub lstAut_ROPDOSISRV_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Integer, K2 As Integer

Me.Enabled = False: Me.MousePointer = vbHourglass
lstAut_ROPDOSISRV_Display
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub lstAut_Usr_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Integer, K2 As Integer
Dim wIndex As Long, xMail As String, xSQL As String

Me.Enabled = False: Me.MousePointer = vbHourglass
oldAut.BIATABID = "ROPDOSGUSR"
K = 0
oldAut.BIATABK1 = Space_Scan(lstAut_Usr.Text, K)
oldAut.BIATABK2 = ""
oldAut.BIATABTXT = ""
If IsNull(sqlYBIATAB0_Read(Trim(oldAut.BIATABID), Trim(oldAut.BIATABK1), Trim(oldAut.BIATABK2), oldAut.BIATABTXT)) Then
    fraAut_Update.Caption = oldAut.BIATABK1 & " : Modification des habilitations "
    cmdAut_SQL_K = "Update"
Else
    fraAut_Update.Caption = oldAut.BIATABK1 & " : Création des habilitations"
    cmdAut_SQL_K = "Insert"
End If
'===========================
newAut = oldAut
'===========================
xSQL = "select SSIDOMUNIT from " & paramIBM_Library_SABSPE & ".YSSIDOM0 where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN'" _
     & " and SSIDOMUIDX = '" & oldAut.BIATABK1 & "'"
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    txtAut_ROPDOSGUSR_SRV = rsSab("SSIDOMUNIT")
    libAut_ROPDOSGUSR_SRV = Trim(arrROPDOSISRV_Code(Val(Mid$(rsSab("SSIDOMUNIT"), 2, 2))))
Else
    txtAut_ROPDOSGUSR_SRV = "S00"
    libAut_ROPDOSGUSR_SRV = ""
    Call MsgBox("Cet utilisateur n'est pas affecté à un service" & vbCrLf & "voir le RSSI pour la mise à jour", vbCritical, "DROPI : habilitations")
End If

txtAut_ROPDOSGUSR_SRV.Locked = True


If Mid$(oldAut.BIATABTXT, 1, 1) = "D" Then
    chkAut_ROPDOSGUSR.Value = "1"
Else
    chkAut_ROPDOSGUSR.Value = "0"
End If
If Mid$(oldAut.BIATABTXT, 2, 1) = "A" Then
    chkAut_ROPINFGUSR.Value = "1"
Else
    chkAut_ROPINFGUSR.Value = "0"
End If
If Mid$(oldAut.BIATABTXT, 3, 1) = "P" Then
    chkAut_ROPDOSGUSR_P.Value = "1"
Else
    chkAut_ROPDOSGUSR_P.Value = "0"
End If
If Mid$(oldAut.BIATABTXT, 4, 1) = "H" Then
    chkAut_ROPDOSGUSR_H.Value = "1"
Else
    chkAut_ROPDOSGUSR_H.Value = "0"
End If
If Mid$(oldAut.BIATABTXT, 5, 1) = "Q" Then
    chkAut_ROPDOSGUSR_Q.Value = "1"
Else
    chkAut_ROPDOSGUSR_Q.Value = "0"
End If
If Mid$(oldAut.BIATABTXT, 6, 1) = "E" Then
    chkAut_ROPDOSGUSR_E.Value = "1"
Else
    chkAut_ROPDOSGUSR_E.Value = "0"
End If
If Mid$(oldAut.BIATABTXT, 7, 1) = "I" Then
    chkAut_ROPDOSGUSR_I.Value = "1"
Else
    chkAut_ROPDOSGUSR_I.Value = "0"
End If

Call lstAut_ROPDOSGUSR_Display(wIndex, txtAut_ROPDOSGUSR_SRV)
If wIndex > 0 Then
    lstAut_ROPDOSISRV.ListIndex = wIndex - 1
    lstAut_ROPDOSISRV_Display
Else
    lstAut_ROPDOSISRV.ListIndex = -1
    optAut_ROPDOSISRV_Z = True
End If
xMail = mailAdresse_Production(oldAut.BIATABK1)
If Trim(xMail) = "" Then
    lblAut_ROPDOSGUSR_Mail.ForeColor = vbRed
    lblAut_ROPDOSGUSR_Mail.Caption = "??? manque adresse mail ??????"
Else
    lblAut_ROPDOSGUSR_Mail.ForeColor = vbBlue
    lblAut_ROPDOSGUSR_Mail.Caption = xMail
End If

fraAut_Update.Visible = True
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub lstParam_ROPDOSISRV_Click()
fraParam_Reset
blnParam_ROPDOSISRV = True
mnuParam_ROPDOSISRV.Visible = True
mnuParam_DOMAINES.Visible = False
mnuparam_APPLICATIONS.Visible = False
mnuparam_LIBELLES.Visible = False
mnuparam_ROPDOSQUAL.Visible = False
mnuparam_ROPDOSQUALB2.Visible = False
txtParam_BIATABK2.Visible = True
oldParam.BIATABID = "ROPDOSISRV"
mlstParam_ListIndex = lstParam_ROPDOSISRV.ListIndex
If lstParam_ROPDOSISRV.ListCount > 0 And lstParam_ROPDOSISRV.ListIndex >= 0 Then
    'kParam_ROPDOSISRV = Val(Mid$(lstParam_ROPDOSISRV.Text, 2, 2))
    oldParam.BIATABK1 = Mid$(lstParam_ROPDOSISRV.Text, 1, 3)
    lstParam_ROPDOSGUSR_Load oldParam.BIATABK1
    '''''Me.PopupMenu mnuparam, vbPopupMenuLeftButton, 800, 8300
Else
    ''''mnuParam_Insert_Click
    
End If

End Sub

Private Sub lstParam_ROPDOSQUAL_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
fraParam_Reset

blnParam_ROPDOSQUAL = True
mnuParam_DOMAINES.Visible = False
mnuparam_APPLICATIONS.Visible = False
mnuparam_LIBELLES.Visible = False
mnuparam_ROPDOSQUAL.Visible = True
mnuparam_ROPDOSQUALB2.Visible = False
mnuParam_ROPDOSISRV.Visible = False
txtParam_BIATABK2.Visible = False

oldParam.BIATABID = "ROPDOSQUAL"
mlstParam_ListIndex = lstParam_ROPDOSQUAL.ListIndex
If lstParam_ROPDOSQUAL.ListCount > 0 And lstParam_ROPDOSQUAL.ListIndex >= 0 Then
    oldParam.BIATABK1 = Mid$(lstParam_ROPDOSQUAL.Text, 1, 3)
 '   oldParam.BIATABK2 = arrROPINFQUAL_BIATABK2(lstParam_ROPINFQUAL.ListIndex)
    mnuParam_Delete.Enabled = True
    Me.PopupMenu mnuparam, vbPopupMenuLeftButton, 8500, 8300
Else
    mnuParam_Insert_Click
    
End If

End Sub





Private Sub lstParam_ROPDOSQUALB2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
fraParam_Reset

blnParam_ROPDOSQUAL = True
mnuParam_DOMAINES.Visible = False
mnuparam_APPLICATIONS.Visible = False
mnuparam_LIBELLES.Visible = False
mnuparam_ROPDOSQUAL.Visible = False
mnuparam_ROPDOSQUALB2.Visible = True
mnuParam_ROPDOSISRV.Visible = False
txtParam_BIATABK2.Visible = False

oldParam.BIATABID = "ROPDOSQUALB2"
mlstParam_ListIndex = lstParam_ROPDOSQUALB2.ListIndex
If lstParam_ROPDOSQUALB2.ListCount > 0 And lstParam_ROPDOSQUALB2.ListIndex >= 0 Then
    oldParam.BIATABK1 = Mid$(lstParam_ROPDOSQUALB2.Text, 1, 1)
 '   oldParam.BIATABK2 = arrROPINFQUAL_BIATABK2(lstParam_ROPINFQUAL.ListIndex)
    mnuParam_Delete.Enabled = True
    Me.PopupMenu mnuparam, vbPopupMenuLeftButton, 800, 8300
Else
    mnuParam_Insert_Click
    
End If

End Sub

Private Sub lstParam_ROPDOSXAPP_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
fraParam_Reset
oldParam.BIATABID = "ROPDOSXAPP"
oldParam.BIATABK1 = Mid$(lstParam_ROPDOSXDOM.Text, 1, 12)
oldParam.BIATABK2 = Mid$(lstParam_ROPDOSXAPP.Text, 1, 12)
blnParam_ROPDOSXDOM = False
mnuParam_DOMAINES.Visible = False
mnuparam_APPLICATIONS.Visible = True
mnuparam_LIBELLES.Visible = False
mnuparam_ROPDOSQUAL.Visible = False
mnuparam_ROPDOSQUALB2.Visible = False
mnuParam_ROPDOSISRV.Visible = False
txtParam_BIATABK2.Visible = True
mlstParam_ListIndex = lstParam_ROPDOSXAPP.ListIndex

If lstParam_ROPDOSXAPP.ListCount > 0 Then
    mnuParam_Delete.Enabled = True
Else
    mnuParam_Delete.Enabled = False
End If
lstParam_ROPDOSXAPP.Enabled = True
Me.PopupMenu mnuparam, vbPopupMenuLeftButton, 5000, 8300

End Sub


Private Sub lstParam_ROPDOSXDOM_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
fraParam_Reset
oldParam.BIATABID = "ROPDOSXDOM"
oldParam.BIATABK1 = Mid$(lstParam_ROPDOSXDOM.Text, 1, 12)

blnParam_ROPDOSXDOM = True
mnuParam_DOMAINES.Visible = True
mnuparam_APPLICATIONS.Visible = False
mnuparam_LIBELLES.Visible = False
mnuparam_ROPDOSQUAL.Visible = False
mnuparam_ROPDOSQUALB2.Visible = False
mnuParam_ROPDOSISRV.Visible = False
txtParam_BIATABK2.Visible = False

lstParam_ROPDOSXAPP_Load oldParam.BIATABK1, 0
lstParam_ROPDOSXAPP.Enabled = True
mlstParam_ListIndex = lstParam_ROPDOSXDOM.ListIndex
If lstParam_ROPDOSXAPP.ListCount > 0 Then
    mnuParam_Delete.Enabled = False
Else
    mnuParam_Delete.Enabled = True
End If

'________________________________________________________________
Dim xDom As String
xDom = Trim(oldParam.BIATABK1)
If xDom = "Incidents si" Or xDom = "Sécurité" Then
    '$JPL 2013-10-01 oldROPDOSMAIL.BIATABID = "ROPDOSMAIL"
    '$JPL 2013-10-01 oldROPDOSMAIL.BIATABK1 = xDom
    '$JPL 2013-10-01 oldROPDOSMAIL.BIATABK2 = ""
    '$JPL 2013-10-01 If IsNull(sqlYBIATAB0_Read(oldROPDOSMAIL.BIATABID, oldROPDOSMAIL.BIATABK1, oldROPDOSMAIL.BIATABK2, oldROPDOSMAIL.BIATABTXT)) Then
     '$JPL 2013-10-01    blnROPDOSMAIL = True
     '$JPL 2013-10-01    txtROPDOSMAIL = Trim(oldROPDOSMAIL.BIATABTXT)
     txtROPDOSMAIL = srvSendMail.Exchange_Distribution("DROPI", xDom)
        libROPDOSMAIL.Visible = True
        txtROPDOSMAIL.Visible = True
        txtROPDOSMAIL.Locked = True
'    End If
End If
'__________________________________________________________________

Me.PopupMenu mnuparam, vbPopupMenuLeftButton, 300, 8300
End Sub


Private Sub lstParam_ROPINFGTXT_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
fraParam_Reset
blnParam_ROPINFGTXT = True
mnuParam_DOMAINES.Visible = False
mnuparam_APPLICATIONS.Visible = False
mnuparam_LIBELLES.Visible = True
mnuparam_ROPDOSQUAL.Visible = False
mnuparam_ROPDOSQUALB2.Visible = False
mnuParam_ROPDOSISRV.Visible = False
txtParam_BIATABK2.Visible = False


oldParam.BIATABID = "ROPINFGTXT"
mlstParam_ListIndex = lstParam_ROPINFGTXT.ListIndex
If lstParam_ROPINFGTXT.ListCount > 0 And lstParam_ROPINFGTXT.ListIndex >= 0 Then
    oldParam.BIATABK1 = Mid$(lstParam_ROPINFGTXT.Text, 1, 12)
    oldParam.BIATABK2 = arrROPINFGTXT_BIATABK2(lstParam_ROPINFGTXT.ListIndex)
    mnuParam_Delete.Enabled = True
    Me.PopupMenu mnuparam, vbPopupMenuLeftButton, 10000, 8300
Else
    mnuParam_Insert_Click
    
End If
End Sub

Private Sub lstUpdate_Modèle_Click()
lstUpdate_Modèle.Visible = False
cmdDossier_Ok.Visible = True: fraDossier_cmd.Visible = False
End Sub

Private Sub lstUpdate_ROPDOSQUAL_Click()
newYROPDOS0 = oldYROPDOS0
newYROPDOS0.ROPDOSQUAL = Mid$(lstUpdate_ROPDOSQUAL, 1, 3)
cmdDossier_Ok_Click
End Sub

Private Sub lstUpdate_ROPINFMAIL_CC_Click()
Dim K As Integer
lstUpdate_ROPINFMAIL_CC_Display.Clear
For K = 0 To lstUpdate_ROPINFMAIL_CC.ListCount - 1

    If lstUpdate_ROPINFMAIL_CC.Selected(K) Then lstUpdate_ROPINFMAIL_CC_Display.AddItem arrROPINFMAIL(K)
Next K

End Sub

Private Sub lstUpdate_ROPINFMAIL_Click()
Dim K As Integer
lstUpdate_ROPINFMAIL_Display.Clear
For K = 0 To lstUpdate_ROPINFMAIL.ListCount - 1

    If lstUpdate_ROPINFMAIL.Selected(K) Then lstUpdate_ROPINFMAIL_Display.AddItem arrROPINFMAIL(K)
Next K
    'For K = 0 To lstUpdate_ROPINFMAIL.ListCount - 1
    '    lstUpdate_ROPINFMAIL.ListIndex = K
    '    If lstUpdate_ROPINFMAIL.Selected(K) Then Call cmdSendMail_Recipient(Trim(lstUpdate_ROPINFMAIL.Text))
    'Next K
    'lstUpdate_ROPINFMAIL.Visible = True: lstUpdate_ROPINFMAIL_Display.Visible = True

End Sub

Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub

Private Sub mnuContextQuitter_Click()
Unload Me
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim meUnit As typeUnit, X As String
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate


mAPP_Menu = Mid$(Msg, 1, 12)
Call BiaPgmAut_Init(mAPP_Menu, DROPI_Aut)

blnSetfocus = True
Form_Init
blnAuto = False

Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case "@AUTO_ROPDOS": blnAuto = True
                        lstSelect_Load_6
                        cmdSelect_SQL_6
                        Unload Me
    Case Else: blnAuto = False
End Select


End Sub


Public Sub cmdContext_Return()
Select Case SSTab1.Tab
    Case Is = 0
        If fgSelect.Visible = False And fraDossier.Visible = False Then cmdSelect_Ok_Click
    Case Else
        If currentAction = "" Then
            If SSTab1.Tab > 0 Then
                SSTab1.Tab = 0
            Else
               'SendKeys "{TAB}"
               ' cmdSelect_Click
            End If
        End If
End Select
End Sub









Public Sub cmdSendMail(lFct As String)
Dim libROPINFGNAT As String
Dim K As Long, K1 As Long, blnDisplay As Boolean
Dim wROPINFSTA As String
Dim mFontFace As String
Dim wROPDOSXID As String, wROPDOSXDOM As String

Dim X1 As String, wUsrName As String
Dim xUsr As String, kLen As Integer, K2 As Integer
Dim meName_Ucase As String
Dim wPJ As String
Dim blnDossierComplet As Boolean


blnSendMail = False
meName_Ucase = usrName_UCase
wRecipient = ""
wccRecipient = ""
arrRecipient_Nb = 0
If lFct = " " Then
    blnDossierComplet = False
Else
    blnDossierComplet = True
End If
'=============================================

If lFct = "X" Then
    lstUpdate_ROPINFMAIL.Visible = False: lstUpdate_ROPINFMAIL_Display.Visible = False:    libUpdate_ROPINFMAIL.Visible = False
    lstUpdate_ROPINFMAIL_CC.Visible = False: lstUpdate_ROPINFMAIL_CC_Display.Visible = False:    libUpdate_ROPINFMAIL_CC.Visible = False
    mailYROPINF0.ROPINFIDP = -1
    
    For K = 0 To lstUpdate_ROPINFMAIL_CC.ListCount - 1
        lstUpdate_ROPINFMAIL_CC.ListIndex = K
        If lstUpdate_ROPINFMAIL_CC.Selected(K) Then Call cmdSendMail_Recipient(Trim(lstUpdate_ROPINFMAIL_CC.Text))
    Next K
    wccRecipient = wRecipient
    wRecipient = ""
    For K = 0 To lstUpdate_ROPINFMAIL.ListCount - 1
        lstUpdate_ROPINFMAIL.ListIndex = K
        If lstUpdate_ROPINFMAIL.Selected(K) Then Call cmdSendMail_Recipient(Trim(lstUpdate_ROPINFMAIL.Text))
    Next K
    lstUpdate_ROPINFMAIL.Visible = True: lstUpdate_ROPINFMAIL_Display.Visible = True:    libUpdate_ROPINFMAIL.Visible = True
    lstUpdate_ROPINFMAIL_CC.Visible = True: lstUpdate_ROPINFMAIL_CC_Display.Visible = True:    libUpdate_ROPINFMAIL_CC.Visible = True
Else
    If lFct = "IncidentSignificatif" Then
        cmdSendMail_ROPDOSMAIL "Incidents si"
    Else
        If lFct = "Sécurité" Then
            cmdSendMail_ROPDOSMAIL "Sécurité"
        Else

'_____________________________________________________________________________________________
            If chkUpdate_ROPINFMAIL <> "1" Then Exit Sub
            '=============================================
            
            For K = 1 To 5
                X1 = Mid$(mailYROPINF0.ROPINFMAIL, K, 1)
                Select Case X1
                    Case "I": wUsrName = Trim(mailYROPDOS0.ROPDOSIUSR)
                    Case "D": wUsrName = Trim(mailYROPDOS0.ROPDOSGUSR)
                    Case "A": wUsrName = Trim(mailYROPINF0.ROPINFGUSR)
                    Case "P": wUsrName = Trim(arrYROPINF0(Processus_Index).ROPINFGUSR)
                    'Case "U": wUsrName = Trim(mailYROPINF0.ROPINFUUSR): meName_Ucase = ""
                    Case "U": wUsrName = Trim(mailYROPINF0_Suivant.ROPINFGUSR): meName_Ucase = ""
                   
                    Case Else: wUsrName = ""
                End Select
                If wUsrName <> "" And wUsrName <> meName_Ucase Then
                    If Mid$(wUsrName, 1, 2) <> "_S" Then
                        Call cmdSendMail_Recipient(wUsrName)
                    Else
                        K1 = Val(Mid$(wUsrName, 3, 2))
                        xUsr = arrROPDOSISRV_Mail(K1)
                        kLen = Len(xUsr)
                        For K2 = 1 To kLen Step 12
                            Call cmdSendMail_Recipient(Mid$(xUsr, K2, 12))
                        Next K2
                        
                    End If
                    
                End If
            Next K
        End If
    End If
End If

'=============================================

If wRecipient = "" Then Exit Sub
'_____________________________________________________________________________________________
Call arrYROPINF0_SQL(mailYROPDOS0.ROPDOSID)
'_____________________________________________________________________________________________
wSendMail.From = currentSSIWINMAIL
wSendMail.FromDisplayName = usrName_UCase
wSendMail.Recipient = wRecipient
wSendMail.CcRecipient = wccRecipient
wSendMail.Attachment = ""

bgColor = "" '"cyan"

'wSendMail.Subject = "RO -" & mailYROPINF0.ROPINFID & "-" & mailYROPINF0.ROPINFIDP & "-" & mailYROPINF0.ROPINFIDT & "-" & mailYROPINF0.ROPINFIDT2 & " : " & mailSubject
wSendMail.Subject = "BIA.RO -" & mailYROPDOS0.ROPDOSID & " : " & Trim(mailYROPDOS0.ROPDOSXDOM) & " - " & Trim(mailYROPDOS0.ROPDOSXAPP) & " - " & Trim(mailYROPDOS0.ROPDOSXID) _
                   & " (" & Trim(fraDossier_Display_USR(mailYROPDOS0.ROPDOSIUSR)) & ")"
'wDétail = "<FONT color=#0000A0  face=" & Asc34 & "Arial" & Asc34 & ">" _
'_____________________________________________________________________________________________
mFontFace = "<FONT face=" & Asc34 & "@Calibri" & Asc34 & ">"
'wDétail = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">" _
'        & mFontFace _

If lFct <> "IncidentSignificatif" Then

    wDétail = "<TR>" _
             & "<TD  bgcolor=#00A0A0 width= 90 height=4><span style='font-size:11.0pt;font-family:Calibri'><Font color=#FFFFFF><b><center>" & mailYROPDOS0.ROPDOSID & "</center></b></TD>" _
             & "<TD  bgcolor=#00A0A0  width= 570 height=4><span style='font-size:11.0pt;font-family:Calibri'><Font color=#FFFFFF><b>" & mailYROPDOS0.ROPDOSXDOM & " - " & mailYROPDOS0.ROPDOSXAPP & "</b></TD>" _
             & "<TD  bgcolor=#00A0A0 width= 210 height=4><span style='font-size:11.0pt;font-family:Calibri'><Font color=#FFFFFF></TD>" _
             & "<TD  bgcolor=#00A0A0 width=30 height=4><span style='font-size:11.0pt;font-family:Calibri'><Font color=#FFFFFF>Etat</TD>" _
            & "</TR>"
Else
    wSendMail.Subject = "Incident significatif " & wSendMail.Subject
    wDétail = "<TR>" _
             & "<TD  bgcolor=#FF0000 width= 90 height=4><span style='font-size:11.0pt;font-family:Calibri'><Font color=#FFFFFF><b><center>" & mailYROPDOS0.ROPDOSID & "</center></b></TD>" _
             & "<TD  bgcolor=#FF0000  width= 570 height=4><span style='font-size:11.0pt;font-family:Calibri'><Font color=#FFFFFF><b>" & mailYROPDOS0.ROPDOSXDOM & " - " & mailYROPDOS0.ROPDOSXAPP & "</b></TD>" _
             & "<TD  bgcolor=#FF0000 width= 210 height=4><span style='font-size:11.0pt;font-family:Calibri'><Font color=#FFFFFF>" & mailYROPDOS0.ROPDOSGCOU & " €</TD>" _
             & "<TD  bgcolor=#FF0000 width=30 height=4><span style='font-size:11.0pt;font-family:Calibri'><Font color=#FFFFFF>Etat</TD>" _
            & "</TR>"
End If

wTD_BackColor = "bgcolor = #00A0A0" '#87CEFA"
wTD_Sta_ForeColor = "<Font color=#FFFFFF>" 'wTD_BackColor
If Trim(mailYROPDOS0.ROPDOSXID) = "" Then
    wROPDOSXID = ""
Else
    wROPDOSXID = " . " & htmlFontColor_Red & mailYROPDOS0.ROPDOSXID
End If

X = cmdSendMail_Sta(mailYROPDOS0.ROPDOSSTA, wTD_Sta_ForeColor)
xUsr = Trim(cmdSendMail_USR(mailYROPDOS0.ROPDOSGUSR))
wDétail = wDétail & "<TR>" _
         & "<TD " & wTD_BackColor & " width= 90 height=4><span style='font-size:10.0pt;font-family:Calibri'><Font color=#FFFFFF>" & dateImp10(mailYROPDOS0.ROPDOSGECH) & "</TD>" _
         & "<TD " & wTD_BackColor & " width= 570 height=4><span style='font-size:10.0pt;font-family:Calibri'><Font color=#FFFFFF><B>" & xUsr & "</B></TD>" _
         & "<TD " & wTD_BackColor & " width= 210 height=4><span style='font-size:9.0pt;font-family:Calibri'><Font color=#FFFFFF>" & dateImp10(mailYROPDOS0.ROPDOSIAMJ) & " - " & Trim(cmdSendMail_USR(mailYROPDOS0.ROPDOSIUSR)) & "</TD>" _
         & "<TD " & wTD_BackColor & " width=30 height=4><span style='font-size:10.0pt;font-family:Calibri'><Font color=#FFFFFF>" & X & "</TD>" _
        & "</TR>"
'_____________________________________________________________________________________________
For K = 1 To arrYROPINF0_Nb
    xYROPINF0 = arrYROPINF0(K)
    blnDisplay = blnDossierComplet
    If xYROPINF0.ROPINFGNAT = "P" Then
        If xYROPINF0.ROPINFIDP = 1 And xYROPINF0.ROPINFIDT = 0 Then
            blnDisplay = True
        Else
            If xYROPINF0.ROPINFIDP = mailYROPINF0.ROPINFIDP Then blnDisplay = True
        End If
    Else
       If xYROPINF0.ROPINFIDP = mailYROPINF0.ROPINFIDP Then blnDisplay = True
    End If
    
    If blnDisplay Then
        If xYROPINF0.ROPINFUAMJ = DSys And xYROPINF0.ROPINFUHMS > 80000 Then
            wTD_Txt_ForeColor = htmlFontColor_Gray 'Blue
        Else
            wTD_Txt_ForeColor = htmlFontColor_Gray
        End If
        If xYROPINF0.ROPINFSTA = " " Then
            If xYROPINF0.ROPINFUAMJ = DSys Then
                wTD_ForeColor = htmlFontColor_Blue
            Else
                wTD_ForeColor = htmlFontColor_Green
            End If
        Else
            wTD_ForeColor = htmlFontColor_Gray
        End If
      
        Select Case xYROPINF0.ROPINFGNAT
            Case "P":
                libROPINFGNAT = xYROPINF0.ROPINFID & " § " & xYROPINF0.ROPINFIDP
                wTD_BackColor = "bgcolor = #FFC080" '#90FFFF"
                wTD_Sta_ForeColor = wTD_BackColor
                X = cmdSendMail_Sta(xYROPINF0.ROPINFSTA, wTD_Sta_ForeColor)
                xUsr = Trim(cmdSendMail_USR(xYROPINF0.ROPINFGUSR))
                wDétail = wDétail & "<TR>" _
                         & "<TD " & wTD_Sta_ForeColor & " width= 90 height=4><span style='font-size:10.0pt;font-family:Calibri'><B>" & cmdSendMail_Ech(xYROPINF0.ROPINFGECH, xYROPINF0.ROPINFSTA) & "</B></TD>" _
                         & "<TD " & wTD_Sta_ForeColor & " width= 570 height=4><span style='font-size:10.0pt;font-family:Calibri'><B>" & wTD_ForeColor & "=> " & xUsr & "</B></TD>" _
                         & "<TD " & wTD_Sta_ForeColor & " width= 210 height=4><span style='font-size:8.0pt;font-family:Calibri'><Font color = #0000FF >" & dateImp10(xYROPINF0.ROPINFCAMJ) & " - " & Trim(cmdSendMail_USR(xYROPINF0.ROPINFCUSR)) & "</TD>" _
                         & "<TD " & wTD_Sta_ForeColor & " width=30 height=4><span style='font-size:10.0pt;font-family:Calibri'>" & X & "</TD>" _
                         & "</TR>" _
                         & "<TR>" _
                         & "<TD colspan=4 width=900 height=4><span style='font-size:11.0pt;font-family:Calibri'>" & wTD_Txt_ForeColor & cmdSendMail_Txt(xYROPINF0.ROPINFGTXT, "H") _
                         & "</TD></TR>"
             Case "A", "F":
                libROPINFGNAT = xYROPINF0.ROPINFID & " § " & xYROPINF0.ROPINFIDP & " - " & Format$(xYROPINF0.ROPINFIDT, "00")
             
                wTD_BackColor = "bgcolor =  #FFCF9B" '#FFE6C8" '#B0FFFF"
                wTD_Sta_ForeColor = wTD_BackColor
                X = cmdSendMail_Sta(xYROPINF0.ROPINFSTA, wTD_Sta_ForeColor)
                'If xYROPINF0.ROPINFIDT = mailYROPINF0_Suivant.ROPINFIDT And mailYROPINF0_Suivant.ROPINFIDT > 0 Then
                '    wTD_Sta_ForeColor = "bgcolor =#FFB000"
                'End If

                xUsr = cmdSendMail_USR(xYROPINF0.ROPINFGUSR)
                wDétail = wDétail & "<TR>" _
                         & "<TD " & wTD_Sta_ForeColor & " width= 90  height=4><span style='font-size:10.0pt;font-family:Calibri'>" & cmdSendMail_Ech(xYROPINF0.ROPINFGECH, xYROPINF0.ROPINFSTA) & "</TD>" _
                         & "<TD " & wTD_Sta_ForeColor & " width= 570  height=4><span style='font-size:10.0pt;font-family:Calibri'><B>" & wTD_ForeColor & "=> " & xUsr & "</B></TD>" _
                         & "<TD " & wTD_Sta_ForeColor & " width= 210  height=4><span style='font-size:8.0pt;font-family:Calibri'><Font color = #0000FF >" & dateImp10(xYROPINF0.ROPINFCAMJ) & " - " & Trim(cmdSendMail_USR(xYROPINF0.ROPINFCUSR)) & "</TD>" _
                         & "<TD " & wTD_Sta_ForeColor & " width=30  height=4><span style='font-size:10.0pt;font-family:Calibri'>" & X & "</TD>" _
                         & "</TR>" _
                         & "<TR>" _
                         & "<TD colspan=4 width=900 height=4><span style='font-size:11.0pt;font-family:Calibri'>" & wTD_Txt_ForeColor & cmdSendMail_Txt(xYROPINF0.ROPINFGTXT, "H") _
                         & "</TD></TR>"
              Case "J"
                    'X = fraDétail_Display_PJ_FileName(xYROPINF0.ROPINFGUSR, False)
                    'wSendMail.Attachment = wSendMail.Attachment & X & ";"
                    wTD_BackColor = "bgcolor = #FFFFDE" '#F5F5F5"
                    wTD_Sta_ForeColor = wTD_BackColor
                    X = cmdSendMail_Sta(xYROPINF0.ROPINFSTA, wTD_Sta_ForeColor)
                    xUsr = cmdSendMail_USR(xYROPINF0.ROPINFGUSR)

                    wPJ = "<a href=" & Asc34 _
                        & paramROPDOS_Path_DROPI & xYROPINF0.ROPINFID _
                        & "\" & xYROPINF0.ROPINFID & "_" & xYROPINF0.ROPINFIDP _
                        & "_" & xYROPINF0.ROPINFIDT & "_" & xYROPINF0.ROPINFIDT2 & "." & UCase$(Trim(xYROPINF0.ROPINFGUSR)) _
                        & Asc34 & ">" & Trim(xYROPINF0.ROPINFGTXT)

                    wDétail = wDétail & "<TR>" _
                         & "<TD " & wTD_Sta_ForeColor & " width= 90 height=4><span style='font-size:8.0pt;font-family:Calibri'>" & wTD_Txt_ForeColor & "PJ" & "</TD>" _
                         & "<TD " & wTD_Sta_ForeColor & " width= 570 height=4><span style='font-size:10.0pt;font-family:Calibri'>" & wTD_Txt_ForeColor & wPJ & "</TD>" _
                         & "<TD " & wTD_Sta_ForeColor & " width= 210 height=4><span style='font-size:8.0pt;font-family:Calibri'>" & wTD_Txt_ForeColor & dateImp10(xYROPINF0.ROPINFCAMJ) & " - " & Trim(cmdSendMail_USR(xYROPINF0.ROPINFCUSR)) & "</TD>" _
                         & "<TD " & wTD_Sta_ForeColor & " width=30 height=4><span style='font-size:10.0pt;font-family:Calibri'>" & X & "</TD>" _
                         & "</TD></TR>"
               Case Else
                    Select Case xYROPINF0.ROPINFGNAT
                        Case "N": libROPINFGNAT = "Note"
                        Case Else: libROPINFGNAT = xYROPINF0.ROPINFGNAT
                    End Select
                    wTD_BackColor = "bgcolor =#FFF0DC" '#F5F5F5"
                    wTD_Sta_ForeColor = wTD_BackColor
                    X = cmdSendMail_Sta(xYROPINF0.ROPINFSTA, wTD_Sta_ForeColor)
                    xUsr = cmdSendMail_USR(xYROPINF0.ROPINFGUSR)
                    wDétail = wDétail & "<TR>" _
                         & "<TD " & wTD_Sta_ForeColor & " width= 90 height=4><span style='font-size:10.0pt;font-family:Calibri'>" & wTD_Txt_ForeColor & libROPINFGNAT & "</TD>" _
                         & "<TD " & wTD_Sta_ForeColor & " width= 570 height=4><span style='font-size:9.0pt;font-family:Calibri'>" & wTD_Txt_ForeColor & xUsr & "</TD>" _
                         & "<TD " & wTD_Sta_ForeColor & " width= 210 height=4><span style='font-size:8.0pt;font-family:Calibri'>" & wTD_Txt_ForeColor & dateImp10(xYROPINF0.ROPINFCAMJ) & " - " & Trim(cmdSendMail_USR(xYROPINF0.ROPINFCUSR)) & "</TD>" _
                         & "<TD " & wTD_Sta_ForeColor & " width=30 height=4><span style='font-size:10.0pt;font-family:Calibri'>" & X & "</TD>" _
                         & "</TR>" _
                         & "<TR>" _
                         & "<TD colspan=4 width=900 height=4><span style='font-size:11.0pt;font-family:Calibri'>" & wTD_Txt_ForeColor & cmdSendMail_Txt(xYROPINF0.ROPINFGTXT, "H") _
                         & "</TD></TR>"
               End Select
            
    End If
Next K
X = "<!DOCTYPE html PUBLIC " & Asc34 & "-//W3C//DTD XHTML 1.0 Transitionnal//EN" & Asc34 & " " & Asc34 & "http://www.w3.org/TR/xhtml1/DTD/transitionnal.dtd" & Asc34 & ">" _
   & "<html>" _
   & "<head>" _
   & "</head>" _
   & "<body>"
   
'<FONT face=" & Asc34 & "@Calibri" & Asc34 & ">
wSendMail.Message = X _
                  & "<span style='font-size:10.0pt;font-family:Calibri'> <TABLE border = 1  width=900 height=4 bgcolor=#FFFFFF cellpadding=4 >" _
        & wDétail _
        & "</TABLE>" _
        & "</body>" _
        & "</html>"

 
wSendMail.AsHTML = True
'___________________________________
'Open "C:\Temp\x.htm" For Output As #99
'Print #99, wSendMail.Message
'Close 99
'MsgBox "cmdsendmail Exit"
'Exit Sub
'___________________________________
If Not blnOff_Line And mailYROPINF0.ROPINFID > 2000 Then
    srvSendMail.Monitor wSendMail
    For K1 = 1 To arrRecipient_Nb
        lstErr.AddItem "@ " & arrRecipient(K1)
    Next K1

End If

End Sub

Public Sub fraDossier_Display_YROPDOS0()
Dim V
Dim X As String, X1 As String
fraDossier.Visible = True
fraDossier_B.Enabled = False
fraDossier_Display_Reset

fraUpdate_B.Enabled = False
Call lstErr_Clear(lstErr, cmdContext, ">Affichage du dossier"): DoEvents
libDossier_ROPDOSID = Trim(xYROPDOS0.ROPDOSID) & " -" & xYROPDOS0.ROPDOSXID
libDossier_ROPDOSIUSR = "=? " & dateImp10(xYROPDOS0.ROPDOSIAMJ) & " " & fraDossier_Display_USR(xYROPDOS0.ROPDOSIUSR)
libDossier_ROPDOSGUSR = "=> " & dateImp10(xYROPDOS0.ROPDOSGECH) & " " & fraDossier_Display_USR(xYROPDOS0.ROPDOSGUSR)
libDossier_ROPDOSUUSR = "Dossier : " & xYROPDOS0.ROPDOSID & "-" & xYROPDOS0.ROPDOSSTA _
                   & "  ( " & Trim(xYROPDOS0.ROPDOSCUSR) & " " & dateImp10(xYROPDOS0.ROPDOSCAMJ) _
                   & " ) ( màj : " & Trim(xYROPDOS0.ROPDOSUUSR) & " " & dateImp10(xYROPDOS0.ROPDOSUAMJ) & " " & timeImp8(xYROPDOS0.ROPDOSUHMS) _
                   & "  v_" & xYROPDOS0.ROPDOSUVER & " )"
cbo_Scan xYROPDOS0.ROPDOSSTA, txtUpdate_ROPDOSSTA


cbo_Scan Trim(xYROPDOS0.ROPDOSXDOM), txtUpdate_ROPDOSXDOM

sqlYBIATAB0_cboID_K1 "ROPDOSXAPP", xYROPDOS0.ROPDOSXDOM, txtUpdate_ROPDOSXAPP
cbo_Scan Trim(xYROPDOS0.ROPDOSXAPP), txtUpdate_ROPDOSXAPP
txtUpdate_ROPDOSIREF = Trim(xYROPDOS0.ROPDOSIREF)
txtUpdate_ROPDOSXID = Trim(xYROPDOS0.ROPDOSXID)
txtUpdate_ROPDOSGCOU = Trim(xYROPDOS0.ROPDOSGCOU)

cbo_Scan Trim(xYROPDOS0.ROPDOSGUSR), txtUpdate_ROPDOSGUSR
If Mid$(xYROPDOS0.ROPDOSGSRV, 1, 2) = "_S" Then libUpdate_ROPDOSGSRV = arrROPDOSISRV_Lib(Val(Mid$(xYROPDOS0.ROPDOSGSRV, 3, 2)))
Call DTPicker_Set(txtUpdate_ROPDOSGECH, xYROPDOS0.ROPDOSGECH)

cbo_Scan Trim(xYROPDOS0.ROPDOSIUSR), txtUpdate_ROPDOSIUSR
If Mid$(xYROPDOS0.ROPDOSISRV, 1, 2) = "_S" Then libUpdate_ROPDOSISRV = arrROPDOSISRV_Lib(Val(Mid$(xYROPDOS0.ROPDOSISRV, 3, 2)))
Call DTPicker_Set(txtUpdate_ROPDOSIAMJ, xYROPDOS0.ROPDOSIAMJ)

cbo_Scan xYROPDOS0.ROPDOSGNAT, txtUpdate_ROPDOSGNAT
cbo_Scan xYROPDOS0.ROPDOSGPRI, txtUpdate_ROPDOSGPRI
cbo_Scan xYROPDOS0.ROPDOSGPRV, txtUpdate_ROPDOSGPRV
cbo_Scan xYROPDOS0.ROPDOSGGRA, txtUpdate_ROPDOSGGRA
'If blnROPDOSQUAL Then
cbo_Scan xYROPDOS0.ROPDOSQUAL, txtUpdate_ROPDOSQUAL

End Sub
Public Function fraDossier_Control()
Dim V, wMsgBox As String
Dim K As Long

Dim blnUpdate_Control As Boolean
Dim X As String
blnUpdate_Control = True
blnIncidentSignificatif_Mail = False
blnSécurité_Mail = False

Call lstErr_AddItem(lstErr, cmdContext, ">Contrôle dossier "): DoEvents
newYROPDOS0 = oldYROPDOS0
fraDossier_Display_Reset

wMsgBox = ""

Call DTPicker_Control(txtUpdate_ROPDOSGECH, newYROPDOS0.ROPDOSGECH)
If cmdUpdate_K = "14" And newYROPDOS0.ROPDOSGECH = oldYROPDOS0.ROPDOSGECH Then
Else
    If newYROPDOS0.ROPDOSGECH < DSys And newYROPDOS0.ROPDOSID > 2000 Then
        lblUpdate_ROPDOSGECH.BackColor = vbRed
        txtUpdate_ROPDOSGECH.ToolTipText = "L'échéance du dossier ne peut pas être < à aujourd'hui " & " | " & dateImp10(arrYROPINF0(K).ROPINFGECH)
        blnUpdate_Control = False
        Call lstErr_AddItem(lstErr, cmdContext, "?_________échéance < " & DSys)
            wMsgBox = wMsgBox & "- L'échéance du dossier ne peut pas être < à aujourd'hui " & vbCrLf
    Else
        For K = 2 To arrYROPINF0_Nb
            If newYROPDOS0.ROPDOSGECH < arrYROPINF0(K).ROPINFGECH Then
                lblUpdate_ROPDOSGECH.BackColor = vbRed
                txtUpdate_ROPDOSGECH.ToolTipText = "L'échéance du dossier ne peut pas être < à l'échéance de l'action " & arrYROPINF0(K).ROPINFIDP & "-" & arrYROPINF0(K).ROPINFIDT & " | " & dateImp10(arrYROPINF0(K).ROPINFGECH)
                blnUpdate_Control = False
                Call lstErr_AddItem(lstErr, cmdContext, "?_________éch Dossier < éch Action " & arrYROPINF0(K).ROPINFIDP & " / " & arrYROPINF0(K).ROPINFIDT)
                wMsgBox = wMsgBox & "- L'échéance du dossier ne peut pas être < à l'échéance de l'action " & arrYROPINF0(K).ROPINFIDP & "-" & arrYROPINF0(K).ROPINFIDT & " | " & dateImp10(arrYROPINF0(K).ROPINFGECH) & vbCrLf
                Exit For
            End If
        Next K
    End If
End If
Call DTPicker_Control(txtUpdate_ROPDOSIAMJ, newYROPDOS0.ROPDOSIAMJ)
If cmdUpdate_K = "14" And newYROPDOS0.ROPDOSIAMJ = oldYROPDOS0.ROPDOSIAMJ Then
Else

    If newYROPDOS0.ROPDOSIAMJ < mROPDOSIAMJ_Min Then
        lblUpdate_ROPDOSIAMJ.BackColor = vbRed
        txtUpdate_ROPDOSIAMJ.ToolTipText = "La date du constat ne peut pas être < à " & dateImp10(mROPDOSIAMJ_Min)
        blnUpdate_Control = False
        Call lstErr_AddItem(lstErr, cmdContext, "?_________date constat erronée")
        wMsgBox = wMsgBox & "- La date du constat ne peut pas être < à " & dateImp10(mROPDOSIAMJ_Min) & vbCrLf
    End If
    If newYROPDOS0.ROPDOSIAMJ > newYROPDOS0.ROPDOSUAMJ Then
        lblUpdate_ROPDOSIAMJ.BackColor = vbRed
        txtUpdate_ROPDOSIAMJ.ToolTipText = "La date du constat ne peut pas être > à " & dateImp10(newYROPDOS0.ROPDOSUAMJ)
        blnUpdate_Control = False
        Call lstErr_AddItem(lstErr, cmdContext, "?_________date constat > date saisie")
        wMsgBox = wMsgBox & "- La date du constat ne peut pas être > à " & dateImp10(newYROPDOS0.ROPDOSUAMJ) & vbCrLf
    End If
End If

newYROPDOS0.ROPDOSGUSR = txtUpdate_ROPDOSGUSR
If Trim(newYROPDOS0.ROPDOSGUSR) = "?" Then
    If Not blnDossierModèle Then
        txtUpdate_ROPDOSGUSR.BackColor = vbRed
        blnUpdate_Control = False
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Gestionnaire ? ")
        wMsgBox = wMsgBox & "- Gestionnaire ? " & vbCrLf
    End If
End If

newYROPDOS0.ROPDOSIUSR = txtUpdate_ROPDOSIUSR
If Trim(newYROPDOS0.ROPDOSIUSR) = "?" Then
    If Not blnDossierModèle Then
        txtUpdate_ROPDOSIUSR.BackColor = vbRed
        blnUpdate_Control = False
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Initiateur du constat ? ")
        wMsgBox = wMsgBox & "- Initiateur du constat ? " & vbCrLf
    End If
End If


newYROPDOS0.ROPDOSGNAT = txtUpdate_ROPDOSGNAT
If Trim(newYROPDOS0.ROPDOSGNAT) = "?" Then
    If Not blnDossierModèle Then
        txtUpdate_ROPDOSGNAT.BackColor = vbRed
        blnUpdate_Control = False
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Nature ? ")
        wMsgBox = wMsgBox & "- Nature ? " & vbCrLf
    End If
End If

newYROPDOS0.ROPDOSGPRV = txtUpdate_ROPDOSGPRV
If Trim(newYROPDOS0.ROPDOSGPRV) = "?" Then
    If Not blnDossierModèle Then
        txtUpdate_ROPDOSGPRV.BackColor = vbRed
        blnUpdate_Control = False
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Confidentialité ? ")
        wMsgBox = wMsgBox & "- Confidentialité ? " & vbCrLf
    End If
End If

newYROPDOS0.ROPDOSGGRA = txtUpdate_ROPDOSGGRA
newYROPDOS0.ROPDOSGPRI = txtUpdate_ROPDOSGPRI
newYROPDOS0.ROPDOSXDOM = Mid$(txtUpdate_ROPDOSXDOM, 1, 12)
If Not blnDossierModèle Then
    If Trim(newYROPDOS0.ROPDOSXDOM) = "?" Then
        txtUpdate_ROPDOSXDOM.BackColor = vbRed
        blnUpdate_Control = False
        Call lstErr_AddItem(lstErr, cmdContext, "?_________préciser le domaine ? ")
        wMsgBox = wMsgBox & "- préciser le domaine ? " & vbCrLf
    Else
        If Trim(newYROPDOS0.ROPDOSXDOM) = "Modèle" Then
            txtUpdate_ROPDOSXDOM.BackColor = vbRed
            txtUpdate_ROPDOSXDOM.ToolTipText = "Le domaine 'Modèle' n'est pas autorisé pour un dossier"
            blnUpdate_Control = False
            Call lstErr_AddItem(lstErr, cmdContext, "?_________'Modèle' n'est pas un domaine autorisé ? ")
            wMsgBox = wMsgBox & "- Le domaine 'Modèle' n'est pas autorisé pour un dossier" & vbCrLf
        End If
    End If
End If
newYROPDOS0.ROPDOSXAPP = Mid$(txtUpdate_ROPDOSXAPP, 1, 12)
newYROPDOS0.ROPDOSIREF = txtUpdate_ROPDOSIREF
newYROPDOS0.ROPDOSXID = txtUpdate_ROPDOSXID
newYROPDOS0.ROPDOSGCOU = Val(txtUpdate_ROPDOSGCOU)
'If blnROPDOSQUAL Then
newYROPDOS0.ROPDOSQUAL = Mid$(txtUpdate_ROPDOSQUAL, 1, 3)

If newYROPDOS0.ROPDOSGPRV = "U" Then

If newYROPDOS0.ROPDOSGNAT = "I" Then
    blnUpdate_Control = False
    Call lstErr_AddItem(lstErr, cmdContext, "? dossier privé 'U' et nature 'I' incompatibles")
    wMsgBox = wMsgBox & "-  dossier privé 'U' et nature 'I' incompatibles" & vbCrLf
End If

    If Trim(newYROPDOS0.ROPDOSGUSR) <> usrName_UCase Then
        txtUpdate_ROPDOSGUSR.BackColor = vbRed
        txtUpdate_ROPDOSGUSR.ToolTipText = "Dossier privé(U) => le superviseur doit être = " & usrName_UCase
        blnUpdate_Control = False
        Call lstErr_AddItem(lstErr, cmdContext, "?_________dossier U => cet utilisateur n'est pas autorisé ")
        wMsgBox = wMsgBox & "- Dossier privé(U) => le superviseur doit être = " & usrName_UCase & vbCrLf
    End If
End If

X = Trim(txtUpdate_ROPDOSGTXT)
X = Replace(X, "<", "{")
X = Replace(X, ">", "}")

If X = "" Then
    txtUpdate_ROPDOSGTXT.BackColor = vbRed
    txtUpdate_ROPDOSGTXT.ToolTipText = "Veuillez préciser un texte  < 1024 caractères"
    blnUpdate_Control = False
    Call lstErr_AddItem(lstErr, cmdContext, "?_________préciser le texte")
    wMsgBox = wMsgBox & "-  Veuillez préciser un texte  < 1024 caractères" & vbCrLf
Else
    txtUpdate_ROPDOSGTXT.BackColor = vbWhite
End If

txtUpdate_ROPINFGUSR = txtUpdate_ROPDOSGUSR
chkUpdate_ROPINFGPRV = "0"
txtUpdate_ROPINFGUO = 0
txtUpdate_ROPINFGTXT = txtUpdate_ROPDOSGTXT
Call DTPicker_Set(txtUpdate_ROPINFGECH, newYROPDOS0.ROPDOSGECH)
'_______________________________________________________________________
If newYROPDOS0.ROPDOSXDOM <> oldYROPDOS0.ROPDOSXDOM Then
    If newYROPDOS0.ROPDOSXDOM = "Incidents si" Then blnIncidentSignificatif_Mail = True
End If
If newYROPDOS0.ROPDOSGCOU <> oldYROPDOS0.ROPDOSGCOU Then
    If newYROPDOS0.ROPDOSGCOU >= 600000 And oldYROPDOS0.ROPDOSGCOU <= 600000 Then
        If newYROPDOS0.ROPDOSXDOM <> "Incidents si" Then blnIncidentSignificatif_Mail = True
    End If
End If
'_______________________________________________________________________
If Trim(newYROPDOS0.ROPDOSXDOM) = "Sécurité" Then
    If newYROPDOS0.ROPDOSXAPP <> oldYROPDOS0.ROPDOSXAPP Then
        If Trim(newYROPDOS0.ROPDOSXAPP) = "personne" Or Trim(newYROPDOS0.ROPDOSXAPP) = "Biens" Then blnSécurité_Mail = True
    End If
End If
'_______________________________________________________________________

If blnUpdate_Control Then
    fraDossier_Control = Null
Else
    fraDossier_Control = "<Fin du contrôle des données "
    Call MsgBox(wMsgBox, vbCritical, "DROPI : nouveau dossier")
End If
End Function

Public Function fraDétail_Update_Control()
Dim V
Dim blnUpdate_Control As Boolean
Dim X As String, lngX As Long, wMsgBox As String
blnUpdate_Control = True
Call lstErr_AddItem(lstErr, cmdContext, ">Contrôle Information "): DoEvents
fraDétail_Display_Reset

wMsgBox = ""
newYROPINF0 = oldYROPINF0
newYROPINF0.ROPINFSTA = " "
newYROPINF0.ROPINFSTAK = " "
newYROPINF0.ROPINFGUSR = txtUpdate_ROPINFGUSR

If Trim(newYROPINF0.ROPINFGUSR) = "?" Then
    If Not blnDossierModèle Then
        txtUpdate_ROPINFGUSR.BackColor = vbRed
        txtUpdate_ROPINFGUSR.ToolTipText = "Préciser le responsable ? "
        blnUpdate_Control = False
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Responsable ? ")
        wMsgBox = wMsgBox & "- Préciser le responsable" & vbCrLf

    End If
End If

newYROPINF0.ROPINFGNAT = txtUpdate_ROPINFGNAT
'newYROPINF0.ROPINFMAIL = txtUpdate_ROPINFMAIL
If chkUpdate_ROPINFMAIL_I = "1" Then
    Mid$(newYROPINF0.ROPINFMAIL, 1, 1) = "I"
Else
    Mid$(newYROPINF0.ROPINFMAIL, 1, 1) = " "
End If
If chkUpdate_ROPINFMAIL_D = "1" Then
    Mid$(newYROPINF0.ROPINFMAIL, 2, 1) = "D"
Else
    Mid$(newYROPINF0.ROPINFMAIL, 2, 1) = " "
End If
If chkUpdate_ROPINFMAIL_P = "1" Then
    Mid$(newYROPINF0.ROPINFMAIL, 3, 1) = "P"
Else
    Mid$(newYROPINF0.ROPINFMAIL, 3, 1) = " "
End If
If chkUpdate_ROPINFMAIL_A = "1" Then
    Mid$(newYROPINF0.ROPINFMAIL, 4, 1) = "A"
Else
    Mid$(newYROPINF0.ROPINFMAIL, 4, 1) = " "
End If
If chkUpdate_ROPINFMAIL_U = "1" Then
    Mid$(newYROPINF0.ROPINFMAIL, 5, 1) = "U"
Else
    Mid$(newYROPINF0.ROPINFMAIL, 5, 1) = " "
End If

If chkUpdate_ROPINFGPRV = "1" Then
    newYROPINF0.ROPINFGPRV = "U"
Else
    newYROPINF0.ROPINFGPRV = " "
End If

Call DTPicker_Control(txtUpdate_ROPINFGECH, newYROPINF0.ROPINFGECH)
If newYROPINF0.ROPINFGNAT = "P" Or newYROPINF0.ROPINFGNAT = "A" Or newYROPINF0.ROPINFGNAT = "F" Then
    If newYROPINF0.ROPINFGECH <> oldYROPINF0.ROPINFGECH Then
        If newYROPINF0.ROPINFGECH < DSys And newYROPINF0.ROPINFID > 2000 Then
            lblUpdate_ROPINFGECH.BackColor = vbRed
            txtUpdate_ROPINFGECH.ToolTipText = "l'échéance ne peut être < à " & dateImp10(DSys)
            blnUpdate_Control = False
            Call lstErr_AddItem(lstErr, cmdContext, "?_________échéance < " & dateImp10(DSys))
            wMsgBox = wMsgBox & "-  l'échéance ne peut être < à " & dateImp10(DSys) & vbCrLf
        End If
        If newYROPINF0.ROPINFGECH > oldYROPDOS0.ROPDOSGECH Then
            lblUpdate_ROPINFGECH.BackColor = vbRed
            txtUpdate_ROPINFGECH.ToolTipText = "l'échéance ne peut être > à l'échéance du dossier " & dateImp10(oldYROPDOS0.ROPDOSGECH)
           blnUpdate_Control = False
            Call lstErr_AddItem(lstErr, cmdContext, "?_________échéance Info > éch Dossier")
            wMsgBox = wMsgBox & "-  l'échéance ne peut être > à l'échéance du dossier " & dateImp10(oldYROPDOS0.ROPDOSGECH) & vbCrLf
        End If
    End If
End If

lngX = CLng(Val(txtUpdate_ROPINFGUO) * 100)
If lngX Mod 100 > 60 Then
    txtUpdate_ROPINFGUO.BackColor = vbRed
    txtUpdate_ROPINFGUO.ToolTipText = "la partie décimale (minute) doit être <= à 60"
    blnUpdate_Control = False
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Durée.Min > 60")
    wMsgBox = wMsgBox & "-  la partie décimale (minute) doit être <= à 60" & vbCrLf
End If
If lngX > 99999 Then
    txtUpdate_ROPINFGUO.BackColor = vbRed
    txtUpdate_ROPINFGUO.ToolTipText = "le nombre d'heures ne peut pas être > à 999 h"
    blnUpdate_Control = False
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Durée.heure > 999")
    wMsgBox = wMsgBox & "-  le nombre d'heures ne peut pas être > à 999 h" & vbCrLf
Else
    newYROPINF0.ROPINFGUO = lngX
End If

If Not blnYROPINF0_12X Then
    X = Trim(txtUpdate_ROPINFGTXT)
    X = Replace(X, "<", "{")
    X = Replace(X, ">", "}")
    
    newYROPINF0.ROPINFGTXT = X
    If X = "" Then
        txtUpdate_ROPINFGTXT.BackColor = vbRed
        txtUpdate_ROPINFGTXT.ToolTipText = "Veuillez préciser un texte  < 1024 caractères"
        blnUpdate_Control = False
        Call lstErr_AddItem(lstErr, cmdContext, "?_________préciser le texte")
    wMsgBox = wMsgBox & "-  Veuillez préciser un texte  < 1024 caractères" & vbCrLf
    End If
End If

If mROPINFSTA_Value <> "" Then newYROPINF0.ROPINFSTA = mROPINFSTA_Value

If blnDossierModèle Then
    newYROPINF0.ROPINFIDTL = Val(txtUpdate_ROPINFIDTL)
   ' If newYROPINF0.ROPINFIDTL = 1 Then
   '     blnUpdate_Control = False
   '     Call lstErr_AddItem(lstErr, cmdContext, "?_________action liée doit être <> 1")
   ' End If
End If

If blnUpdate_Control Then
    fraDétail_Update_Control = Null
Else
    fraDétail_Update_Control = "<Fin du contrôle des données "
    Call MsgBox(wMsgBox, vbCritical, "DROPI : contrôle des informations")
End If
End Function

Public Sub fraDétail_Display()
Dim V, I As Integer, K As Integer
Dim X As String, X1 As String

blnControl = False
fraDétail_Display_Reset


libDossier_ROPINFID = "dossier : " & xYROPINF0.ROPINFID & "   " & xYROPINF0.ROPINFIDP _
                   & "-" & xYROPINF0.ROPINFIDT & "-" & xYROPINF0.ROPINFIDT2 & "-" & xYROPINF0.ROPINFSTA & "-" & xYROPINF0.ROPINFSTAD _
                   & "  ( " & Trim(xYROPINF0.ROPINFCUSR) & " " & dateImpS(xYROPINF0.ROPINFCAMJ) _
                   & " ) ( màj : " & Trim(xYROPINF0.ROPINFUUSR) & " " & dateImpS(xYROPINF0.ROPINFUAMJ) & " " & timeImp8(xYROPINF0.ROPINFUHMS) _
                   & "  v_" & xYROPINF0.ROPINFUVER & " )"


fraDétail_Update_Enabled

Call lstErr_Clear(lstErr, cmdContext, ">Affichage du détail"): DoEvents


cbo_Scan xYROPINF0.ROPINFSTA, txtUpdate_ROPINFSTA

cbo_Scan Trim(xYROPINF0.ROPINFGUSR), txtUpdate_ROPINFGUSR
'________________________________________________________________________________
If Mid$(xYROPINF0.ROPINFMAIL, 2, 1) = "D" Then
    chkUpdate_ROPINFMAIL_D = "1"
Else
    chkUpdate_ROPINFMAIL_D = "0"
End If
chkUpdate_ROPINFMAIL_D.Caption = fraDossier_Display_USR(oldYROPDOS0.ROPDOSGUSR)

If Mid$(xYROPINF0.ROPINFMAIL, 1, 1) = "I" Then
    chkUpdate_ROPINFMAIL_I = "1"
Else
    chkUpdate_ROPINFMAIL_I = "0"
End If
chkUpdate_ROPINFMAIL_I.Caption = fraDossier_Display_USR(oldYROPDOS0.ROPDOSIUSR)

If Mid$(xYROPINF0.ROPINFMAIL, 3, 1) = "P" Then
    chkUpdate_ROPINFMAIL_P = "1"
Else
    chkUpdate_ROPINFMAIL_P = "0"
End If
If Processus_Index > 0 Then chkUpdate_ROPINFMAIL_P.Caption = fraDossier_Display_USR(arrYROPINF0(Processus_Index).ROPINFGUSR)

If Mid$(xYROPINF0.ROPINFMAIL, 4, 1) = "A" Then
    chkUpdate_ROPINFMAIL_A = "1"
Else
    chkUpdate_ROPINFMAIL_A = "0"
End If
If Action_Index > 0 Then
    chkUpdate_ROPINFMAIL_A.Caption = fraDossier_Display_USR(xYROPINF0.ROPINFGUSR)
Else
    chkUpdate_ROPINFMAIL_A.Caption = ""
End If

chkUpdate_ROPINFMAIL_U = "0"
If Action_Suivante_Index > 0 Then
    chkUpdate_ROPINFMAIL_U.Caption = fraDossier_Display_USR(arrYROPINF0(Action_Suivante_Index).ROPINFGUSR)
    If xYROPINF0.ROPINFGECH <= YBIATAB0_DATE_CPT_JS1 Then chkUpdate_ROPINFMAIL_U = "1"
Else
    chkUpdate_ROPINFMAIL_U.Caption = ""
End If
'chkUpdate_ROPINFMAIL_U.Caption = usrName_UCase
'_______________________________________________________________________________________________________________
If chkUpdate_ROPINFMAIL = "1" Then fraUpdate_ROPINFMAIL.Visible = True

If xYROPINF0.ROPINFGPRV = "U" Then
    chkUpdate_ROPINFGPRV = "1"
Else
    chkUpdate_ROPINFGPRV = "0"
End If


Call DTPicker_Set(txtUpdate_ROPINFGECH, xYROPINF0.ROPINFGECH)
cbo_Scan xYROPINF0.ROPINFGNAT, txtUpdate_ROPINFGNAT
If xYROPINF0.ROPINFUVER = 0 And oldYROPDOS0.ROPDOSGNAT <> "M" Then
   txtUpdate_ROPINFGTXT_0.Visible = True
   txtUpdate_ROPINFGTXT_0 = Trim(oldYROPINF0.ROPINFGTXT)
   txtUpdate_ROPINFGTXT = ""
Else
    txtUpdate_ROPINFGTXT_0.Visible = False
    txtUpdate_ROPINFGTXT = Trim(xYROPINF0.ROPINFGTXT)
End If
If xYROPINF0.ROPINFGUO <> 0 Then
    txtUpdate_ROPINFGUO = Format$(xYROPINF0.ROPINFGUO / 100, "##0.00")
Else
    txtUpdate_ROPINFGUO = ""
End If
If xYROPINF0.ROPINFGNAT = "J" Then
' click droit     X = fraDétail_Display_PJ_FileName(xYROPINF0.ROPINFGUSR, True)

End If

cmdUpdate.Enabled = True
blnROPINFIDTL_Ok = True

If blnDossierModèle Then
    lblUpdate_ROPINFIDTL.Visible = True
    txtUpdate_ROPINFIDTL.Visible = True
    txtUpdate_ROPINFIDTL.Enabled = True
    txtUpdate_ROPINFIDTL = xYROPINF0.ROPINFIDTL
Else
    lblUpdate_ROPINFIDTL.Visible = False
    txtUpdate_ROPINFIDTL.Visible = False
    If xYROPINF0.ROPINFIDTL <> 0 Then
        For I = 1 To arrYROPINF0_Nb
            If arrYROPINF0(I).ROPINFIDP = xYROPINF0.ROPINFIDP _
            And arrYROPINF0(I).ROPINFIDT = xYROPINF0.ROPINFIDTL _
            And arrYROPINF0(I).ROPINFIDT2 = 1 Then
                If arrYROPINF0(I).ROPINFSTA = " " Then
                    blnROPINFIDTL_Ok = False
                    lblUpdate_ROPINFIDTL.Visible = True
                    txtUpdate_ROPINFIDTL.Visible = True
                    txtUpdate_ROPINFIDTL.Enabled = False
                    txtUpdate_ROPINFIDTL = xYROPINF0.ROPINFIDTL
                        
                    Exit For
                End If
            End If
        Next I
    End If
End If
'fgDetail.LeftCol = 0
blnControl = True

End Sub

Private Sub cmdUpdate_Init_04()
cmdUpdate_K = "04"
currentAction = "cmdUpdate_Init_04"

blnControl = False
mailSubject = "Création d'un dossier"
'fraUpdate_B.Enabled = True

'txtUpdate_ROPINFGUSR.Enabled = True
'txtUpdate_ROPINFGNAT.Enabled = False
'txtUpdate_ROPINFGECH.Enabled = True 'False
'If oldYROPDOS0.ROPDOSID >= 2000 Then
'    chkUpdate_ROPINFMAIL_I = "0"
'    chkUpdate_ROPINFMAIL_D = "0"
'    chkUpdate_ROPINFMAIL_P = "0"
'    chkUpdate_ROPINFMAIL_A = "0"
'    chkUpdate_ROPINFMAIL_U = "0"
'End If
'chkUpdate_ROPINFGPRV = "0"
'txtUpdate_ROPINFGTXT.Locked = False

'txtUpdate_ROPINFGNAT.Enabled = False
'txtUpdate_ROPINFGUO = "": txtUpdate_ROPINFGUO.Enabled = False
'Call fraDétail_lbl("P", 1)

'fraDossier_B.Visible = True
'blnSendMail = True

'fraUpdate_B.Visible = True
'tabDossier.Tab = 1: tabDossier.Caption = "Suivi"
tabDossier.Tab = 0
oldYROPINF0 = currentYROPINF0

'libUpdate_ROPDOSGSRV = ""
'libUpdate_ROPDOSISRV = ""
txtUpdate_ROPDOSGTXT = ""
'cmdUpdate_Dossier.Visible = False
'cmdUpdate.Visible = False
fgDetail.Visible = False
txtUpdate_ROPDOSGTXT.Locked = False
blnControl = True
End Sub

Private Sub cmdUpdate_Init_11()
mailSubject = "Modification d'une information"
fraUpdate_B.Enabled = True
txtUpdate_ROPINFGUSR.Enabled = True
txtUpdate_ROPINFGNAT.Enabled = False
txtUpdate_ROPINFGECH.Enabled = True
txtUpdate_ROPINFGUO.Enabled = True

If blnYROPINF0_12X Then
    txtUpdate_ROPINFGTXT.Locked = True
    libDossier_ROPINFGTXT.Visible = False
Else
    txtUpdate_ROPINFGTXT.Locked = False
    libDossier_ROPINFGTXT.Visible = True
End If

cmdDossier_Ok.Visible = True: fraDossier_cmd.Visible = False
On Error Resume Next
txtUpdate_ROPINFGTXT.SetFocus
txtUpdate_ROPINFGTXT.SelStart = Len(txtUpdate_ROPINFGTXT)

End Sub

Private Sub cmdUpdate_Init_01()

cmdUpdate_K = "01"

blnControl = False
mailSubject = "Ajouter une note"
fraUpdate_B.Enabled = True
txtUpdate_ROPINFGUSR.Enabled = True
txtUpdate_ROPINFGNAT.Enabled = False
txtUpdate_ROPINFGECH.Enabled = False
chkUpdate_ROPINFGPRV = "1"

cbo_Scan Trim(usrName_UCase), txtUpdate_ROPINFGUSR
cbo_Scan "N", txtUpdate_ROPINFGNAT
'cbo_Scan " ", txtUpdate_ROPINFMAIL
Call DTPicker_Set(txtUpdate_ROPINFGECH, DSys)
txtUpdate_ROPINFGTXT = ""

txtUpdate_ROPINFGTXT.Locked = False
cmdDossier_Ok.Visible = True: fraDossier_cmd.Visible = False
txtUpdate_ROPINFGUO = "": txtUpdate_ROPINFGUO.Enabled = False
Call fraDétail_lbl("N", 0)

cmdUpdate.Visible = False
fraUpdate_B.Visible = True
tabDossier.Tab = 2
tabDossier.Caption = mailSubject
oldYROPINF0 = currentYROPINF0

blnControl = True
End Sub


Private Sub cmdUpdate_Init_05()

cmdUpdate_K = "05"

blnControl = False
mailSubject = "Ajouter une pièce jointe"

rtfPJ.Top = 3720
rtfPJ.Left = 3240
rtfPJ.Width = 5000
rtfPJ.Height = 2100

fraUpdate_B.Enabled = True
txtUpdate_ROPINFGUSR.Enabled = False
txtUpdate_ROPINFGNAT.Enabled = False
txtUpdate_ROPINFGECH.Enabled = False
'txtUpdate_ROPINFMAIL.Enabled = True

''''cbo_Scan Trim(usrName_UCase), txtUpdate_ROPINFGUSR
cbo_Scan "J", txtUpdate_ROPINFGNAT
'cbo_Scan " ", txtUpdate_ROPINFMAIL
Call DTPicker_Set(txtUpdate_ROPINFGECH, DSys)
txtUpdate_ROPINFGTXT = ""

txtUpdate_ROPINFGTXT.Locked = True 'False
fraUpdate_PJ.Visible = True
filDoc.Pattern = "_.*"
filDoc.Pattern = "*.*"

cmdDossier_Ok.Visible = True: fraDossier_cmd.Visible = False
txtUpdate_ROPINFGUO = "": txtUpdate_ROPINFGUO.Enabled = False
oldFileName = "": newFileName = ""
rtfPJ.Text = ""

tabDossier.Tab = 3
tabDossier.Caption = mailSubject
oldYROPINF0 = currentYROPINF0

blnControl = True
End Sub

Private Sub cmdUpdate_Init_02(lROPINFGNAT As String)
cmdUpdate_K = "02"
blnControl = False
mailSubject = "Ajouter une action"
fraUpdate_B.Enabled = True
txtUpdate_ROPINFGUSR.Enabled = True
txtUpdate_ROPINFGNAT.Enabled = False
txtUpdate_ROPINFGECH.Enabled = True
txtUpdate_ROPINFGUO.Enabled = True
chkUpdate_ROPINFGPRV = "0" '"1"
    
'If oldYROPDOS0.ROPDOSGPRV = "U" Then
    cbo_Scan Trim(usrName_UCase), txtUpdate_ROPINFGUSR
'Else
'    txtUpdate_ROPINFGUSR.ListIndex = 0
'End If
cbo_Scan lROPINFGNAT, txtUpdate_ROPINFGNAT
Call DTPicker_Set(txtUpdate_ROPINFGECH, oldYROPDOS0.ROPDOSGECH)  'arrYROPINF0(Processus_Index).ROPINFGECH)
txtUpdate_ROPINFGTXT = ""

txtUpdate_ROPINFGTXT.Locked = False
cmdDossier_Ok.Visible = True: fraDossier_cmd.Visible = False
cmdDossier_Ok_Close.Visible = False
txtUpdate_ROPINFGUO = "": txtUpdate_ROPINFGUO.Enabled = True
Call fraDétail_lbl("A", 0)

libDossier_ROPINFGTXT.Visible = True


cmdUpdate.Visible = False
fraUpdate_B.Visible = True
tabDossier.Tab = 2
tabDossier.Caption = mailSubject
oldYROPINF0 = currentYROPINF0

blnControl = True
End Sub
Private Sub cmdUpdate_Init_03()
cmdUpdate_K = "03"
blnControl = False
mailSubject = "Création d'un processus"
fraUpdate_B.Enabled = True
txtUpdate_ROPINFGUSR.Enabled = True
txtUpdate_ROPINFGNAT.Enabled = False
txtUpdate_ROPINFGECH.Enabled = True
chkUpdate_ROPINFGPRV = "1"

If oldYROPDOS0.ROPDOSGPRV = "U" Then
    cbo_Scan Trim(usrName_UCase), txtUpdate_ROPINFGUSR
Else
    txtUpdate_ROPINFGUSR.ListIndex = 0
End If
cbo_Scan "P", txtUpdate_ROPINFGNAT
Call DTPicker_Set(txtUpdate_ROPINFGECH, oldYROPDOS0.ROPDOSGECH)
txtUpdate_ROPINFGTXT = ""

txtUpdate_ROPINFGTXT.Locked = False
cmdDossier_Ok.Visible = True: fraDossier_cmd.Visible = False
txtUpdate_ROPINFGUO = "": txtUpdate_ROPINFGUO.Enabled = False
Call fraDétail_lbl("P", 0)


cmdUpdate.Visible = False
fraUpdate_B.Visible = True
tabDossier.Tab = 2
oldYROPINF0 = currentYROPINF0

blnControl = True
End Sub

Private Sub cmdUpdate_Init_14()
mailSubject = "Modification d'un dossier"
cmdUpdate_Dossier.Visible = False
fraDossier_B.Enabled = True
cmdDossier_Ok.Visible = True: fraDossier_cmd.Visible = False
blnSelect_Update_B_Display = True
End Sub

Private Sub cmdUpdate_Init_14Q()
mailSubject = "Modification de la qualification d'un dossier"
lstUpdate_ROPDOSQUAL.Visible = True
End Sub

Private Sub cmdUpdate_Init_31()
Dim X As String
mailSubject = "Annulation d'une note"
X = "Voulez-vous réellement ANNULER cette information?"
mROPINFSTA_Value = "A"
X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, libDossier_ROPINFID.Caption)
If X = vbYes Then
    cmdDétail_Update_Ok
Else
    cmdUpdate.ListIndex = -1
End If
End Sub

Private Sub cmdUpdate_Init_21()
Dim X As String
mailSubject = "Fermeture d'une note"
X = "Voulez-vous réellement Clôturer cette information?"
mROPINFSTA_Value = "F"
X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraUpdate_B)
If X = vbYes Then
    cmdDétail_Update_Ok
Else
    cmdUpdate.ListIndex = -1
End If

End Sub

Private Sub cmdUpdate_Init_51()
Dim X As String
mailSubject = "Restauration d'une note"

X = "Voulez-vous réellement RESTAURER cette information?"
mROPINFSTA_Value = " "
X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraUpdate_B)
If X = vbYes Then
    cmdDétail_Update_Ok
Else
    cmdUpdate.ListIndex = -1
End If

End Sub

Private Sub cmdUpdate_Init_52()
Dim X As String
mailSubject = "Restauration d'une action"

X = "Voulez-vous réellement RESTAURER cette action?"
mROPINFSTA_Value = " ": mROPINFSTA_Set = " ":: mROPINFSTAK_Set = " "
If oldYROPINF0.ROPINFSTA = "F" Then
    mROPINFSTA_Where = "+"
Else
    mROPINFSTA_Where = "-"
End If


X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraUpdate_B)
If X = vbYes Then
    cmdDétail_Update_Ok
Else
    cmdUpdate.ListIndex = -1
End If
End Sub

Private Sub cmdUpdate_Init_53()
Dim X As String
mailSubject = "Restauration d'un processus"

X = "Voulez-vous réellement RESTAURER ce processus?"
mROPINFSTA_Value = " ": mROPINFSTA_Set = " ": mROPINFSTAK_Set = " "
If oldYROPINF0.ROPINFSTA = "F" Then
    mROPINFSTA_Where = "F"
Else
    mROPINFSTA_Where = "A"
End If


X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraUpdate_B)
If X = vbYes Then
    cmdDétail_Update_Ok

'=============================================
    Call arrYROPINF0_SQL(oldYROPDOS0.ROPDOSID)
    xYROPDOS0 = oldYROPDOS0
    fraDossier_STAK
    
    newYROPDOS0 = oldYROPDOS0
    newYROPDOS0.ROPDOSSTA = " " 'mROPINFSTA_Value
    newYROPDOS0.ROPDOSSTAK = xYROPDOS0.ROPDOSSTAK
    cmdDossier_Ok_Transaction "Update"
'=============================
Else
    cmdUpdate.ListIndex = -1
End If
End Sub


Private Sub cmdUpdate_Init_22()
Dim X As String
mailSubject = "Fermeture d'une action"

X = "Voulez-vous réellement Clôturer cette Action?"
mROPINFSTA_Value = "F": mROPINFSTA_Set = "+": mROPINFSTA_Where = " ": mROPINFSTAK_Set = "V"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraUpdate_B)
If X = vbYes Then
    If Mid$(cmdUpdate, 4, 1) = "+" Then    ' fermeture de l'action et création d'une autre action
        cmdDétail_Update_Ok
        arrSelect_Update_Nb = 1
        arrSelect_Update(1) = "02 - Ajouter une action"
        cmdUpdate_Display (arrSelect_Update(1))
        'cmdUpdate.Clear
        'cmdUpdate.AddItem "- 02 - Ajouter une action"
        cmdUpdate.ListIndex = 0
    Else
       ' cmdDétail_Update_Ok
        blnSendMail = True
        If oldYROPINF0.ROPINFGNAT <> "F" Then
            cmdUpdate_K = "22"
            cmdDétail_Update_Ok
        Else
            cmdUpdate_K = "23"
            cmdUpdate_Init_23_Ok
        End If

    End If
Else
    cmdUpdate.ListIndex = -1

End If


End Sub

Private Sub cmdUpdate_Init_23()
Dim X As String
mailSubject = "Fermeture d'un processus"

X = "Voulez-vous réellement Clôturer ce processus?"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraUpdate_B)
If X = vbYes Then
    cmdUpdate_Init_23_Ok
Else
    cmdUpdate.ListIndex = -1
End If

End Sub
Private Sub cmdUpdate_Init_24()
Dim X As String
mailSubject = "Fermeture d'un dossier"

X = "Voulez-vous réellement Clôturer ce dossier?"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPDOSGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraUpdate_B)
If X = vbYes Then
    newYROPDOS0 = oldYROPDOS0
    cmdUpdate_Init_24_Ok
Else
    cmdUpdate_Dossier.ListIndex = -1
End If
End Sub

Private Sub cmdUpdate_Init_54()
Dim X As String
mailSubject = "Réactivation d'un dossier"

X = "Voulez-vous réellement REACTIVER ce dossier?"
mROPINFSTA_Value = " ": mROPINFSTA_Set = " ": mROPINFSTAK_Set = " "
mROPINFSTAD_Set = " "
If oldYROPDOS0.ROPDOSSTA = "F" Then
    mROPINFSTAD_Where = "F"
Else
    mROPINFSTAD_Where = "A"
End If

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraUpdate_B)
If X = vbYes Then
    Call arrYROPINF0_SQL(oldYROPDOS0.ROPDOSID)
    xYROPDOS0 = oldYROPDOS0
    fraDossier_STAK
    
    newYROPDOS0 = oldYROPDOS0
    newYROPDOS0.ROPDOSSTA = " " 'mROPINFSTA_Value
    newYROPDOS0.ROPDOSSTAK = xYROPDOS0.ROPDOSSTAK
    cmdDossier_Ok_Transaction "Update"
Else
    cmdUpdate_Dossier.ListIndex = -1
End If
End Sub

Private Sub cmdUpdate_Init_64()
Dim X As String
Dim V, K As Long
Dim wNb As Long
mailSubject = "Report d'échéance d'un dossier"

X = InputBox("Indiquer le nombre de jours de report (0 pour abandonner)?")
If Not IsNumeric(X) Then
    MsgBox ("abandon : saisie non valide")
    cmdUpdate_Dossier.ListIndex = -1
    Exit Sub
End If
wNb = Val(X)

If wNb = 0 Then
    MsgBox ("abandon")
    Exit Sub
End If
'___________________________________________________________________________________
Call arrYROPINF0_SQL(oldYROPDOS0.ROPDOSID)

newYROPDOS0 = oldYROPDOS0
newYROPDOS0.ROPDOSGECH = DateAdd_AMJ("d", wNb, oldYROPDOS0.ROPDOSGECH)

'-------------------------------------------------------
App_Debug = "cmdUpdate_Init_64"
'-------------------------------------------------------

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If IsNull(V) Then
'________________________________________________________________________________
    For K = 1 To arrYROPINF0_Nb
        If arrYROPINF0(K).ROPINFSTA = " " Then
           newYROPINF0 = arrYROPINF0(K)
           newYROPINF0.ROPINFGECH = DateAdd_AMJ("d", wNb, arrYROPINF0(K).ROPINFGECH)
           V = sqlYROPINF0_Update(newYROPINF0, arrYROPINF0(K), True)
           If Not IsNull(V) Then Exit For
        End If
    Next K
End If
'________________________________________________________________________________
If Not IsNull(V) Then
    V = cnSAB_Transaction("Rollback")
Else
    V = cnSAB_Transaction("Commit")
    cmdDossier_Ok_Transaction ("Update")
End If
    

'------------------------------------------
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Sub
Private Sub cmdUpdate_Init_74()
mailSubject = "Duplication d'un modèle"
cmdUpdate_Dossier.Visible = False
fraUpdate_B.Enabled = True
cmdDossier_Ok.Visible = False: fraDossier_cmd.Visible = False
lstUpdate_Modèle.Visible = True
End Sub

Private Sub cmdUpdate_Init_34()
Dim X As String
mailSubject = "Annulation d'un dossier"

X = "Voulez-vous réellement ANNULER ce dossier?"
'mROPINFSTA_Value = "A": mROPINFSTA_Set = "£": mROPINFSTA_Where = " ": mROPINFSTAK_Set = "A"
mROPINFSTAD_Set = "A": mROPINFSTAK_Set = "A": mROPINFSTAD_Where = " "

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPDOSGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraUpdate_B)
If X = vbYes Then
    newYROPDOS0 = oldYROPDOS0
    newYROPDOS0.ROPDOSSTA = "A"
    newYROPDOS0.ROPDOSSTAK = "A"
    cmdDossier_Ok_Transaction "Update"
Else
    cmdUpdate_Dossier.ListIndex = -1
End If
End Sub

Private Sub cmdUpdate_Init_32()
Dim X As String
mailSubject = "Annulation d'une action"

X = "Voulez-vous réellement ANNULER cette Action?"
mROPINFSTA_Value = "A": mROPINFSTA_Set = "-": mROPINFSTA_Where = " ": mROPINFSTAK_Set = "A"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraUpdate_B)
If X = vbYes Then
    cmdDétail_Update_Ok
Else
    cmdUpdate.ListIndex = -1
End If
End Sub

Private Sub cmdUpdate_Init_33()
Dim X As String
mailSubject = "Annulation d'un processus"

X = "Voulez-vous réellement ANNULER ce processus?"
mROPINFSTA_Value = "A": mROPINFSTA_Set = "A": mROPINFSTA_Where = " ": mROPINFSTAK_Set = "A"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraUpdate_B)
If X = vbYes Then
    cmdDétail_Update_Ok
Else
    cmdUpdate.ListIndex = -1
End If
End Sub

Private Sub cmdUpdate_Init_42()
Dim X As String
mailSubject = "Effacement d'une action"

X = "Voulez-vous EFFACER défitivement cette Action?"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraUpdate_B)
If X = vbYes Then
    mROPINFSTA_Set = "E"
    cmdUpdate_Fct = "Delete"
    cmdDétail_Update_Ok
Else
    cmdUpdate.ListIndex = -1
End If
End Sub

Private Sub cmdUpdate_Init_43()
Dim X As String
mailSubject = "Effacement d'un processus"

X = "Voulez-vous EFFACER défitivement ce Processus?"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraUpdate_B)
If X = vbYes Then
    mROPINFSTA_Set = "E"
    cmdUpdate_Fct = "Delete"
    cmdDétail_Update_Ok
Else
    cmdUpdate.ListIndex = -1
End If
End Sub
Private Sub cmdUpdate_Init_44()
Dim X As String
mailSubject = "Effacement d'un dossier"

X = "Voulez-vous EFFACER défitivement ce dossier ?"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPDOSGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraUpdate_B)
If X = vbYes Then
    cmdDossier_Ok_Transaction "Delete"
    cmdSelect_Reset
    Call cmdSelect_SQL_1(0)
Else
    cmdUpdate_Dossier.ListIndex = -1
End If
End Sub

Private Sub cmdUpdate_Init_41()
Dim X As String, K As Integer
mailSubject = "Effacement d'une note"
    
X = "Voulez-vous  EFFACER défitivement cette information?"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraUpdate_B)
If X = vbYes Then
    mROPINFSTA_Set = "E"
    cmdUpdate_Fct = "Delete"
    cmdDétail_Update_Ok
Else
    cmdUpdate.ListIndex = -1
End If
End Sub





Private Sub mnuExport_Click()
fraExport.Visible = True


End Sub

Private Sub mnuExport_Migration_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Export par service ......"): DoEvents

YROPDOS0_Xls_Migration

fraExport.Visible = False

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuExport_Param_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Export param en cours ......"): DoEvents

lstParam_Export

fraExport.Visible = False

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuExport_Service_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Export par service ......"): DoEvents

YROPDOS0_Xls_Service

fraExport.Visible = False

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuParam_Copy_Click()
Call sqlYBIATAB0_Read(Trim(oldParam.BIATABID), Trim(oldParam.BIATABK1), Trim(oldParam.BIATABK2), oldParam.BIATABTXT)
fraParam_Update.Caption = "Copie d'un enregistrement"
cmdParam_SQL_K = "Insert"
fraParam_Display
txtParam_BIATABTXT.Enabled = True
If Trim(oldParam.BIATABID) = "ROPDOSXAPP" Then
    txtParam_BIATABK2.Enabled = True
Else
    txtParam_BIATABK1.Enabled = True
End If

End Sub

Private Sub mnuParam_Delete_Click()
If IsNull(sqlYBIATAB0_Read(Trim(oldParam.BIATABID), Trim(oldParam.BIATABK1), Trim(oldParam.BIATABK2), oldParam.BIATABTXT)) Then
    fraParam_Update.Caption = "Suppression de l'enregistrement"
    cmdParam_SQL_K = "Delete"
    fraParam_Display
End If

End Sub

Private Sub mnuParam_Insert_Click()
fraParam_Update.Caption = "Création d'un enregistrement"
cmdParam_SQL_K = "Insert"
oldParam.BIATABK2 = ""
fraParam_Display
txtParam_BIATABTXT.Enabled = True
If Trim(oldParam.BIATABID) = "ROPDOSXAPP" Then
    txtParam_BIATABK2.Enabled = True
Else
    txtParam_BIATABK1.Enabled = True
End If

End Sub

Private Sub mnuParam_Update_Click()
If IsNull(sqlYBIATAB0_Read(Trim(oldParam.BIATABID), Trim(oldParam.BIATABK1), Trim(oldParam.BIATABK2), oldParam.BIATABTXT)) Then
    fraParam_Update.Caption = "Modification de l'enregistrement"
    fraParam_Update.Caption = "Modification de l'enregistrement"
    cmdParam_SQL_K = "Update"
    fraParam_Display
    txtParam_BIATABTXT.Enabled = True
End If
End Sub

Private Sub mnuPrint0_Dossier_All_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint0_Dossier_All
Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint0_Dossier_All_XDOM_Click()
Dim K As Integer, X As String

Me.Enabled = False: Me.MousePointer = vbHourglass
lstW.Clear
For K = 1 To selYROPDOS0_Nb
    X = Space$(34)
    Mid$(X, 1) = selYROPDOS0(K).ROPDOSXDOM
    Mid$(X, 13) = selYROPDOS0(K).ROPDOSXAPP
    Mid$(X, 25) = Format$(selYROPDOS0(K).ROPDOSID, "0000000000")
    
    lstW.AddItem X
Next K

cmdPrint0_Dossier_All_XDOM
Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub optAut_ROPDOSISRV_C_Click()
If Me.Enabled Then fraAut_Control

End Sub

Private Sub optAut_ROPDOSISRV_D_Click()
If Me.Enabled Then fraAut_Control

End Sub

Private Sub optAut_ROPDOSISRV_H_Click()
If Me.Enabled Then fraAut_Control

End Sub

Private Sub optAut_ROPDOSISRV_I_Click()
If Me.Enabled Then fraAut_Control

End Sub

Private Sub optAut_ROPDOSISRV_R_Click()
If Me.Enabled Then fraAut_Control

End Sub

Private Sub optAut_ROPDOSISRV_X_Click()
If Me.Enabled Then fraAut_Control

End Sub

Private Sub optAut_ROPDOSISRV_Z_Click()
If Me.Enabled Then fraAut_Control

End Sub

Private Sub rtfPJ_Click()
rtfPJ.Top = 240
rtfPJ.Left = 120
rtfPJ.Width = 8500
rtfPJ.Height = 6670
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 4 Then SSTab1.Tab = 0

End Sub

Private Sub SSTab2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
fraParam_Reset
blnParam_ROPINFGTXT = False
mnuParam_DOMAINES.Visible = False
mnuparam_APPLICATIONS.Visible = False
mnuparam_LIBELLES.Visible = True
mnuparam_ROPDOSQUAL.Visible = False
mnuparam_ROPDOSQUALB2.Visible = False
fraParam_Update.Visible = False
End Sub


Private Sub fraDossier_Display()
Dim lenX As Integer, K As Integer
Dim X As String, xSQL As String
Dim xKey As String, xParent As String
Dim blnDossier_Click As Boolean
Me.Enabled = False

tabDossier.Tab = 2
tabDossier.Caption = ""
tabDossier.Tab = 3
tabDossier.Caption = ""

tabDossier.Tab = 1
tabDossier.Caption = "Suivi du dossier : " & xYROPDOS0.ROPDOSID
'tabDossier.ForeColor = &H4000&  ' RGB(0, 0, 240) 'RGB(128, 0, 128)

Select Case xYROPDOS0.ROPDOSSTA
    Case "F": fraUpdate_B.Caption = "Dossier clôturé : " & xYROPDOS0.ROPDOSID
    Case "A": fraUpdate_B.Caption = "Dossier annulé : " & xYROPDOS0.ROPDOSID
    Case Else: fraUpdate_B.Caption = "Dossier  : " & xYROPDOS0.ROPDOSID

End Select

libDossier_ROPINFGTXT.Visible = False
lstUpdate_ROPINFMAIL.Visible = False: lstUpdate_ROPINFMAIL_Display.Visible = False: libUpdate_ROPINFMAIL.Visible = False
lstUpdate_ROPINFMAIL_CC.Visible = False: lstUpdate_ROPINFMAIL_CC_Display.Visible = False: libUpdate_ROPINFMAIL_CC.Visible = False

cmdUpdate_Dossier.Visible = False
fraUpdate_B.Visible = False
fraUpdate_PJ.Visible = False
cmdDossier_Ok.Visible = True: fraDossier_cmd.Visible = False
cmdDossier_Ok_Close.Visible = True

Select Case xYROPDOS0.ROPDOSSTA
    Case "F", "A": fraDossier.BackColor = RGB(164, 164, 164): fraDossier_C.BackColor = RGB(230, 230, 230)
    Case Else: fraDossier.BackColor = RGB(0, 128, 128): fraDossier_C.BackColor = &HE0FFE0

End Select
Call usrColor_Container(fraDossier_C, fraDossier_C.BackColor)

fraDossier_Display_YROPDOS0
oldYROPDOS0 = xYROPDOS0

arrYROPINF0_SQL xYROPDOS0.ROPDOSID
fraDossier_Display_YROPINF0



'____________________________________
fraDossier.Visible = True
fraDossier_B.Visible = True
txtUpdate_ROPDOSGTXT.Locked = True
cmdUpdate_Init_K = "D"
cmdUpdate_Dossier.Clear

cmdUpdate_Init
fraDossier_cmd.Visible = True

Me.Enabled = True
If cmdDossier_Quit.Enabled Then cmdDossier_Quit.SetFocus
End Sub

Public Function fraDossier_Display_USR(lK1 As String) As String
Dim K As Integer, X As String

If Mid$(lK1, 1, 1) <> "_" Then
    fraDossier_Display_USR = Format(lK1, "@")
Else
    K = Val(Mid$(lK1, 3, 2))
    fraDossier_Display_USR = Format(arrROPDOSISRV_Code(K), "@")
End If
End Function

Public Function fgSelect_Display_USR(lK1 As String) As String
Dim K As Integer, X As String

If Mid$(lK1, 1, 1) <> "_" Then
    fgSelect_Display_USR = Format(lK1, "@")
Else
    K = Val(Mid$(lK1, 3, 2))
    fgSelect_Display_USR = Format(arrROPDOSISRV_Code(K), "@")
End If
End Function

Public Function fraDossier_Display_USR_Name(lK1 As String) As String
Dim K As Integer, X As String

If Mid$(lK1, 1, 1) <> "_" Then
    fraDossier_Display_USR_Name = lK1
Else
    K = Val(Mid$(lK1, 3, 2))
    fraDossier_Display_USR_Name = arrROPDOSISRV_Lib(K)
End If
End Function

Private Sub fraDossier_STAK()

Dim kProcessus As Long, K As Long
'___________________________________________________________________________________
kProcessus = 0
blnProcessus_EnCours = False: blnProcessus_EnAlerte = False
blnAction_EnCours = False: blnAction_EnAlerte = False: blnAction_Valide = False
        
'______________________________________________________________________ Modèle
If xYROPDOS0.ROPDOSID < 1000 Then
    xYROPDOS0.ROPDOSSTAK = "M"
    For K = 1 To arrYROPINF0_Nb
        arrYROPINF0(K).ROPINFSTAK = "B"
    Next K
    Exit Sub
End If
'____________________________________________________________________________
For K = 1 To arrYROPINF0_Nb
    xYROPINF0 = arrYROPINF0(K)
    Select Case xYROPINF0.ROPINFGNAT
'____________________________________________________________________________Action
        Case "A", "F":
            Select Case xYROPINF0.ROPINFSTA
                Case "F": arrYROPINF0(K).ROPINFSTAK = "V": blnAction_Valide = True
                Case "A": arrYROPINF0(K).ROPINFSTAK = "A"
                Case " ":
                    blnAction_EnCours = True: blnAction_Valide = True
                    If xYROPINF0.ROPINFGECH > DSys Then
                        arrYROPINF0(K).ROPINFSTAK = "B"
                    Else
                        blnAction_EnAlerte = True
                        If xYROPINF0.ROPINFGECH = DSys Then
                            arrYROPINF0(K).ROPINFSTAK = "O"
                        Else
                            arrYROPINF0(K).ROPINFSTAK = "R"
                        End If
                    End If
                Case Else: arrYROPINF0(K).ROPINFSTAK = " "
            End Select
 '____________________________________________________________________________Processus
        Case "P":
            Call fraDossier_STAK_Processus(kProcessus)
            kProcessus = K
 '____________________________________________________________________________Autres
        Case Else:
            If xYROPINF0.ROPINFSTA = "A" Then
                arrYROPINF0(K).ROPINFSTAK = "A"
            Else
                arrYROPINF0(K).ROPINFSTAK = " "
            End If
    End Select
Next K
'_________________________________________________________________________Dossier

Call fraDossier_STAK_Processus(kProcessus)

Select Case xYROPDOS0.ROPDOSSTA
    Case "F": xYROPDOS0.ROPDOSSTAK = "V"
    Case "A": xYROPDOS0.ROPDOSSTAK = "A"
    Case " ":
        If xYROPDOS0.ROPDOSGECH > DSys Then
            If blnProcessus_EnAlerte Then
                xYROPDOS0.ROPDOSSTAK = "!"
            Else
                If blnProcessus_EnCours Then
                    xYROPDOS0.ROPDOSSTAK = "B"
                Else
                    xYROPDOS0.ROPDOSSTAK = "V"
                    xYROPDOS0.ROPDOSSTA = "F"
                End If
            End If
        Else
            If blnProcessus_EnAlerte Then
                If xYROPDOS0.ROPDOSGECH = DSys Then
                    xYROPDOS0.ROPDOSSTAK = "O"
                Else
                    xYROPDOS0.ROPDOSSTAK = "R"
                End If
            Else
                If blnProcessus_EnCours Then
                    xYROPDOS0.ROPDOSSTAK = "R"
                Else
                    xYROPDOS0.ROPDOSSTAK = "V"
                    xYROPDOS0.ROPDOSSTA = "F"
                End If
            End If
        End If
Case Else:
    xYROPDOS0.ROPDOSSTAK = " "
End Select
'__________________________________________________________

End Sub

Private Sub tabDossier_Click(PreviousTab As Integer)
cmdUpdate.Visible = False
cmdUpdate_Dossier.Visible = False

Select Case tabDossier.Tab
    Case Is = 0:    cmdUpdate_Dossier.Visible = Not cmdDossier_Ok.Visible
    Case 2:    cmdUpdate.Visible = fraUpdate_B.Visible
End Select

End Sub


Private Sub tabDossier_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If cmdSelect_SQL_K <> "2" Then
    If y < 300 Then
        Select Case X
            
            Case Is < 2300: If tabDossier.Tab <> 0 Then tabDossier.Tab = 0
            Case Is < 4600: If tabDossier.Tab <> 1 Then tabDossier.Tab = 1
            Case Is < 6900: If fraUpdate_B.Visible And tabDossier.Tab <> 2 Then tabDossier.Tab = 2
            Case Else:
                    If fraUpdate_PJ.Visible Or lstUpdate_ROPINFMAIL.Visible Then
                            If tabDossier.Tab <> 3 Then tabDossier.Tab = 3
                       End If
        End Select
    End If
End If
End Sub

Private Sub txtAut_ROPDOSGUSR_SRV_Change()
If IsNumeric(Mid$(txtAut_ROPDOSGUSR_SRV, 2, 2)) Then
    libAut_ROPDOSGUSR_SRV = Trim(arrROPDOSISRV_Code(Val(Mid$(txtAut_ROPDOSGUSR_SRV, 2, 2))))
Else
    libAut_ROPDOSGUSR_SRV = ""
End If

End Sub

Private Sub txtAut_ROPDOSGUSR_SRV_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtROPDOSMAIL_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtSelect_ROPDOSGECH_Max_GotFocus()
cmdSelect_Reset

End Sub


Private Sub txtSelect_ROPDOSGNAT_GotFocus()
cmdSelect_Reset

End Sub


Private Sub txtSelect_ROPDOSGPRV_GotFocus()
cmdSelect_Reset

End Sub


Private Sub txtSelect_ROPDOSGUSR_GotFocus()
cmdSelect_Reset
txt_GotFocus txtSelect_ROPDOSGUSR

End Sub

Private Sub txtSelect_ROPDOSGUSR_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub



Public Sub cmdSelect_SQL_6()
Dim xWhere As String
Dim K As Long
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_6"): DoEvents

xWhere = " where ROPDOSSTA = ' ' order by ROPDOSID"
Call arrYROPDOS0_SQL(xWhere)

For arrYROPDOS0_Index = 1 To arrYROPDOS0_Nb
    oldYROPDOS0 = arrYROPDOS0(arrYROPDOS0_Index)
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "Màj STAK : " & oldYROPDOS0.ROPDOSID): DoEvents
    xYROPDOS0 = oldYROPDOS0
    Call arrYROPINF0_SQL(oldYROPDOS0.ROPDOSID)
'_________________________________________________________________________
    ReDim selYROPINF0(arrYROPINF0_Nb)
        For K = 1 To arrYROPINF0_Nb
            selYROPINF0(K) = arrYROPINF0(K)
        Next K
    '_________________________________________________________________________
        cmdUpdate_K = 0
        fraDossier_STAK
'_________________________________________________________________________
    cmdSelect_SQL_6_Transaction
'_________________________________________________________________________

Next arrYROPDOS0_Index

End Sub
 
Public Sub cmdSelect_SQL_JPL()
Dim V, xSQL As String
Dim xWhere As String
Dim K As Long

'=========================================
Exit Sub
'=========================================


Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_JPL"): DoEvents
'20101021 - inversion ROPDOSIREF et ROPINFXID
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then Exit Sub

'________________________________________________________________________________

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 where ROPDOSID >= 2997 order by ROPDOSID"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYROPDOS0_GetBuffer(rsSab, oldYROPDOS0)

     If IsNull(V) Then
        If Trim(oldYROPDOS0.ROPDOSIREF) <> "" Or Trim(oldYROPDOS0.ROPDOSXID) <> "" Then
            Call lstErr_ChangeLastItem(lstErr, cmdContext, "Màj DOS : " & oldYROPDOS0.ROPDOSID): DoEvents
            newYROPDOS0 = oldYROPDOS0
            newYROPDOS0.ROPDOSIREF = oldYROPDOS0.ROPDOSXID
            newYROPDOS0.ROPDOSXID = oldYROPDOS0.ROPDOSIREF
            V = sqlYROPDOS0_Update(newYROPDOS0, oldYROPDOS0, False)
        End If
    End If
    rsSab.MoveNext

Loop
V = cnSAB_Transaction("Commit")

'=========================================
Exit Sub
'=========================================

'20071029 - ajout ROPDOSGSRV et ROPINFGSRV
'_______________________________________________
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then Exit Sub

'________________________________________________________________________________

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0  order by ROPDOSID"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYROPDOS0_GetBuffer(rsSab, oldYROPDOS0)

     If IsNull(V) Then
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "Màj DOS : " & oldYROPDOS0.ROPDOSID): DoEvents
        newYROPDOS0 = oldYROPDOS0
        Call cmdDossier_Ok_GSRV(newYROPDOS0.ROPDOSIUSR, newYROPDOS0.ROPDOSISRV)
        Call cmdDossier_Ok_GSRV(newYROPDOS0.ROPDOSGUSR, newYROPDOS0.ROPDOSGSRV)
        V = sqlYROPDOS0_Update(newYROPDOS0, oldYROPDOS0, False)
    End If
    rsSab.MoveNext

Loop
V = cnSAB_Transaction("Commit")
'_________________________________________________________________________
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then Exit Sub

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YROPINF0" _
     & " order by ROPINFID,ROPINFIDP,ROPINFIDT,ROPINFIDT2"
     
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYROPINF0_GetBuffer(rsSab, oldYROPINF0)

     If IsNull(V) Then
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "Màj INF : " & oldYROPINF0.ROPINFID): DoEvents
        newYROPINF0 = oldYROPINF0
        If oldYROPINF0.ROPINFGNAT = "J" Or oldYROPINF0.ROPINFGNAT = "M" Then
        Else
            Call cmdDossier_Ok_GSRV(newYROPINF0.ROPINFGUSR, newYROPINF0.ROPINFGSRV)
            V = sqlYROPINF0_Update(newYROPINF0, oldYROPINF0, False)
        End If

    End If
    rsSab.MoveNext

Loop
'_________________________________________________________________________
V = cnSAB_Transaction("Commit")

End Sub

Public Sub cmdSelect_SQL_S1()
Dim V, xSQL As String
Dim xWhere As String
Dim Nb As Long, nbD As Long, nbD_F As Long

Call lstErr_Clear(lstErr, cmdContext, "Export > C:\Temp\DROPI_S1.txt"): DoEvents
Call FEU_ROUGE
Open "C:\Temp\DROPI_S1.txt" For Output As #2
rsYROPDOS0_Init oldYROPDOS0
nbD = 0: nbD_F = 0
'________________________________________________________________________________

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0" _
    & " where ROPDOSID >= 1000 and ROPDOSGNAT = 'I'" _
    & " and ropdosiamj > '20070000' and ropdosiamj < '20079999'" _
    & " order by ROPDOSQUAL , ROPDOSXDOM , ROPDOSXAPP "
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYROPDOS0_GetBuffer(rsSab, xYROPDOS0)

     If IsNull(V) Then
        Nb = Nb + 1
        Debug.Print xYROPDOS0.ROPDOSID; xYROPDOS0.ROPDOSQUAL; xYROPDOS0.ROPDOSXDOM, xYROPDOS0.ROPDOSXAPP
        If oldYROPDOS0.ROPDOSQUAL <> xYROPDOS0.ROPDOSQUAL _
        Or oldYROPDOS0.ROPDOSXDOM <> xYROPDOS0.ROPDOSXDOM _
        Or oldYROPDOS0.ROPDOSXAPP <> xYROPDOS0.ROPDOSXAPP Then
            If nbD > 0 Then
                X = oldYROPDOS0.ROPDOSQUAL & ";" & oldYROPDOS0.ROPDOSXDOM & ";" & oldYROPDOS0.ROPDOSXAPP _
                 & ";" & nbD & ";" & nbD_F
                Print #2, X
            End If
            oldYROPDOS0 = xYROPDOS0
            nbD = 0: nbD_F = 0
        End If
        nbD = nbD + 1
        If xYROPDOS0.ROPDOSSTA <> " " Then nbD_F = nbD_F + 1
        
    End If
    rsSab.MoveNext

Loop
If nbD > 0 Then
    X = oldYROPDOS0.ROPDOSQUAL & ";" & oldYROPDOS0.ROPDOSXDOM & ";" & oldYROPDOS0.ROPDOSXAPP _
     & ";" & nbD & ";" & nbD_F
    Print #2, X
End If
Close #2
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "nb  dossiers : " & Nb): DoEvents


End Sub

Public Sub cmdSelect_SQL_7()
Dim V, Nb As Long, nbD As Long, wROPINFID As Long
Dim xWhere As String, xAnd As String, xSQL As String
Dim wIndex As String, X As String
Dim blnOk As Boolean, blnROPDOSXDOM As Boolean, blnROPDOSXAPP As Boolean

mDestinataire_Select = Trim(txtSelect_ROPDOSGUSR)
If mDestinataire_Select = "" Then
    blnDestinataire_Select = False
Else
    blnDestinataire_Select = True
End If

Call DTPicker_Control(txtSelect_ROPDOSGECH_Max, WAMJMax)
meYROPDOS0.ROPDOSXDOM = Mid$(txtSelect_ROPDOSXDOM, 1, 12)
If Trim(meYROPDOS0.ROPDOSXDOM) = "" Then
    blnROPDOSXDOM = False
Else
    blnROPDOSXDOM = True
End If

meYROPDOS0.ROPDOSXAPP = Mid$(txtSelect_ROPDOSXAPP, 1, 12)
If Trim(meYROPDOS0.ROPDOSXAPP) = "" Then
    blnROPDOSXAPP = False
Else
    blnROPDOSXAPP = True
End If
    
xWhere = " where ROPINFSTA = ' ' and ROPINFID > 1000 and  ROPINFIDT > 0 and  ROPINFIDT2 = 1 and ROPINFGECH <= '" & WAMJMax & "'"

xSQL = "select  count(*)  as Tally from " & paramIBM_Library_SABSPE & ".YROPINF0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)
Nb = rsSab("Tally")
ReDim arrYROPINF0(Nb + 10)

arrYROPINF0_Nb = 0: nbD = 0: wROPINFID = 0
xSQL = "select *  from " & paramIBM_Library_SABSPE & ".YROPINF0 " & xWhere _
     & " order by ROPINFID , ROPINFIDP , ROPINFIDT"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    V = rsYROPINF0_GetBuffer(rsSab, xYROPINF0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmROPINF.cmdSelect_SQL_7"
     Else
         arrYROPINF0_Nb = arrYROPINF0_Nb + 1
         arrYROPINF0(arrYROPINF0_Nb) = xYROPINF0
         If wROPINFID <> xYROPINF0.ROPINFID Then
            wROPINFID = xYROPINF0.ROPINFID
            nbD = nbD + 1
        End If
    End If
    rsSab.MoveNext
Loop

ReDim arrYROPDOS0(nbD + 10)
lstW.Clear

arrYROPDOS0_Nb = 0: wROPINFID = 0
blnOk = False

For arrYROPINF0_Index = 1 To arrYROPINF0_Nb
    xYROPINF0 = arrYROPINF0(arrYROPINF0_Index)
    If wROPINFID <> xYROPINF0.ROPINFID Then
        wROPINFID = xYROPINF0.ROPINFID
        xSQL = "select *  from " & paramIBM_Library_SABSPE & ".YROPDOS0 where ROPDOSID = " & xYROPINF0.ROPINFID
        Set rsSab = cnsab.Execute(xSQL)
        If rsSab.EOF Then
            rsYROPDOS0_Init xYROPDOS0
             MsgBox V, vbCritical, "frmDROPI.cmdSelect_SQL_7.2"
        Else
            V = rsYROPDOS0_GetBuffer(rsSab, xYROPDOS0)
        End If
         If Not IsNull(V) Then
             MsgBox V, vbCritical, "frmDROPI.cmdSelect_SQL_7.3"
         Else
            blnOk = True
            If blnROPDOSXDOM Then
                If meYROPDOS0.ROPDOSXDOM <> xYROPDOS0.ROPDOSXDOM Then blnOk = False
                If blnROPDOSXAPP Then
                    If meYROPDOS0.ROPDOSXAPP <> xYROPDOS0.ROPDOSXAPP Then blnOk = False
                End If
            End If
            
            If blnOk Then
                arrYROPDOS0_Nb = arrYROPDOS0_Nb + 1
                arrYROPDOS0(arrYROPDOS0_Nb) = xYROPDOS0
            End If
        End If
    End If
    
    If blnOk Then
        wIndex = "_" & Format$(xYROPINF0.ROPINFID, "000000000") _
                & "_" & Format$(xYROPINF0.ROPINFIDP, "000000000") _
                & "_" & Format$(xYROPINF0.ROPINFIDT, "000000000") _
                & "_" & Format$(arrYROPINF0_Index, "000000000") _
                & "_" & Format$(arrYROPDOS0_Nb, "000000000") _
                & "_" & xYROPINF0.ROPINFGECH _
                & "_" & xYROPINF0.ROPINFGUSR
        
        Call cmdSelect_SQL_7_Destinataire(wIndex)
    End If
            
Next arrYROPINF0_Index
lstW.Visible = True

If cmdSelect_SQL_K = "7@" Then
    cmdSendMail_Echéancier
Else
    cmdPrint0_Echéancier
End If

Exit Sub

End Sub

Private Sub txtSelect_ROPDOSGUSR_LostFocus()
txt_LostFocus txtSelect_ROPDOSGUSR

End Sub

Private Sub txtSelect_ROPDOSID_GotFocus()
cmdSelect_Reset
txt_GotFocus txtSelect_ROPDOSID
End Sub


Private Sub txtSelect_ROPDOSID_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub

Private Sub txtSelect_ROPDOSID_LostFocus()
txt_LostFocus txtSelect_ROPDOSID

End Sub


Private Sub txtSelect_ROPDOSSTA_GotFocus()
cmdSelect_Reset
txt_GotFocus txtSelect_ROPDOSSTA

End Sub


Private Sub txtSelect_ROPDOSSTA_LostFocus()
txt_LostFocus txtSelect_ROPDOSSTA

End Sub


Private Sub txtSelect_ROPDOSUAMJ_Change()
cmdSelect_Reset

End Sub

Private Sub txtSelect_ROPDOSUAMJ_GotFocus()
cmdSelect_Reset

End Sub


Private Sub txtSelect_ROPDOSXAPP_GotFocus()
cmdSelect_Reset
txt_GotFocus txtSelect_ROPDOSXAPP

End Sub


Private Sub txtSelect_ROPDOSXAPP_LostFocus()
txt_LostFocus txtSelect_ROPDOSXAPP

End Sub


Private Sub txtSelect_ROPDOSXDOM_Click()
cmdSelect_Reset
If txtSelect_ROPDOSXDOM.ListIndex = 0 Then
    txtSelect_ROPDOSXAPP.Clear
Else
    sqlYBIATAB0_cboID_K1 "ROPDOSXAPP", Mid$(txtSelect_ROPDOSXDOM, 1, 12), txtSelect_ROPDOSXAPP
    txtSelect_ROPDOSXAPP.AddItem "             - toutes les applications"
    txtSelect_ROPDOSXAPP.ListIndex = 0
End If
End Sub

Private Sub txtSelect_ROPDOSXDOM_GotFocus()
txt_GotFocus txtSelect_ROPDOSXDOM

End Sub


Private Sub txtSelect_ROPDOSXDOM_LostFocus()
txt_LostFocus txtSelect_ROPDOSXDOM

End Sub


Private Sub txtSelect_ROPDOSXID_GotFocus()
cmdSelect_Reset
txt_GotFocus txtSelect_ROPDOSXID

End Sub


Private Sub txtSelect_ROPDOSXID_LostFocus()
txt_LostFocus txtSelect_ROPDOSXID
End Sub


Private Sub txtSelect_ROPINFGTXT_GotFocus()
cmdSelect_Reset
txt_GotFocus txtSelect_ROPINFGTXT


End Sub


Private Sub txtSelect_ROPINFGTXT_LostFocus()
txt_LostFocus txtSelect_ROPINFGTXT

End Sub


Private Sub txtUpdate_ROPDOSGCOU_GotFocus()
txt_GotFocus txtUpdate_ROPDOSGCOU

End Sub

Private Sub txtUpdate_ROPDOSGCOU_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtUpdate_ROPDOSGCOU_LostFocus()
txt_LostFocus txtUpdate_ROPDOSGCOU
End Sub

Private Sub txtUpdate_ROPDOSGECH_Change()
If cmdSelect_SQL_K = "2" _
And Not blntxtUpdate_ROPINFGECH_Change Then txtUpdate_ROPINFGECH = txtUpdate_ROPDOSGECH

End Sub


Private Sub txtUpdate_ROPDOSGECH_GotFocus()
lblUpdate_ROPDOSGECH.BackColor = focusUsr.BackColor
End Sub


Private Sub txtUpdate_ROPDOSGECH_LostFocus()
lblUpdate_ROPDOSGECH.BackColor = fraDossier_C.BackColor '&HC0F0FF

End Sub


Private Sub txtUpdate_ROPDOSGGRA_GotFocus()
txt_GotFocus txtUpdate_ROPDOSGGRA


End Sub


Private Sub txtUpdate_ROPDOSGGRA_LostFocus()
txt_LostFocus txtUpdate_ROPDOSGGRA

End Sub


Private Sub txtUpdate_ROPDOSGNAT_GotFocus()
txt_GotFocus txtUpdate_ROPDOSGNAT

End Sub


Private Sub txtUpdate_ROPDOSGNAT_LostFocus()
txt_LostFocus txtUpdate_ROPDOSGNAT

End Sub


Private Sub txtUpdate_ROPDOSGPRI_GotFocus()
txt_GotFocus txtUpdate_ROPDOSGPRI

End Sub


Private Sub txtUpdate_ROPDOSGPRI_LostFocus()
txt_LostFocus txtUpdate_ROPDOSGPRI

End Sub


Private Sub txtUpdate_ROPDOSGPRV_Click()
If blnControl Then
    If Mid$(txtUpdate_ROPDOSGPRV, 1, 1) = "U" Then
        chkUpdate_ROPINFMAIL_I = "0"
        chkUpdate_ROPINFMAIL_D = "0"
        chkUpdate_ROPINFMAIL_P = "0"
        chkUpdate_ROPINFMAIL_A = "0"
        chkUpdate_ROPINFMAIL_U = "0"
    End If
End If
End Sub

Public Sub fraDétail_Update_Enabled()
Dim blnDétail_Update_INF As Boolean
fraUpdate_B.Enabled = False

txtUpdate_ROPINFGTXT.Locked = True
'cmdDétail_Update_Ok.Visible = False

Call fraDétail_lbl(oldYROPINF0.ROPINFGNAT, oldYROPINF0.ROPINFIDP)

blnDétail_Update_INF = False


If oldYROPDOS0.ROPDOSUUSR = usrName_UCase Or oldYROPINF0.ROPINFUUSR = usrName_UCase Then
    blnDétail_Update_INF = DROPI_Aut.Saisir
Else
    If DROPI_Aut.Xspécial Then
        blnDétail_Update_INF = True
    End If
End If


End Sub


Public Sub fgSelect_YROPDOS0_Forecolor()
fgSelect.Col = 0
Select Case xYROPDOS0.ROPDOSSTA
    Case " ":    fgSelect.CellForeColor = RGB(64, 64, 128) '&HF00000
    Case "A":    fgSelect.CellForeColor = RGB(200, 200, 200) '&H808080
    Case Else: fgSelect.CellForeColor = &H8000&
End Select

Select Case xYROPDOS0.ROPDOSSTAK
    Case "V": fgSelect.CellForeColor = RGB(0, 80, 0) '"Vert"
    Case "B": fgSelect.CellForeColor = RGB(0, 0, 150) '"Bleu"
    Case "O": fgSelect.CellForeColor = RGB(240, 120, 0) ' "Orange"
    Case "R": fgSelect.CellForeColor = RGB(225, 0, 0) '"Rouge"
    Case "A": fgSelect.CellForeColor = RGB(128, 128, 128) '"Non"
    Case "!": fgSelect.CellForeColor = RGB(255, 0, 255) '"Attention"
    Case "M": fgSelect.CellForeColor = RGB(128, 0, 128) '"GrandStroumpf"
    Case Else: fgSelect.CellForeColor = RGB(0, 0, 0) '"D"
End Select
End Sub

Public Sub cmdUpdate_Init()
Dim blnROPINFGPRV As Boolean
Dim wUsr_Aut As String
Dim arrSelect_Update_K As Integer
Dim K As Integer, I As Integer
Dim blnDossierSAnsModèle As Boolean
Dim blnProcessusAvecActionFinale As Boolean

arrSelect_Update_K = 1
fraDossier_cmd.Visible = False

cmdUpdate.Clear
cmdUpdate_12.Visible = False
cmdUpdate_22.Visible = False
cmdUpdate_32.Visible = False

libDossier_ROPINFGTXT.Visible = False
lstUpdate_ROPINFMAIL.Visible = False: lstUpdate_ROPINFMAIL_Display.Visible = False: libUpdate_ROPINFMAIL.Visible = False
lstUpdate_ROPINFMAIL_CC.Visible = False: lstUpdate_ROPINFMAIL_CC_Display.Visible = False: libUpdate_ROPINFMAIL_CC.Visible = False
blnYROPINF0_12X_Aut = False
blnROPINFIDT_Insérer = False
cmdUpdate_Fct = "Update"
cmdDossier_Ok.Caption = "Enregistrer"
mROPINFSTA_Value = "": mROPINFSTA_Set = "": mROPINFSTA_Where = ""
'________________________________________________________________ habilitations
If xYROPINF0.ROPINFGPRV = "U" Then
    If usrName_UCase = Trim(xYROPINF0.ROPINFUUSR) _
    Or usrName_UCase = Trim(xYROPINF0.ROPINFGUSR) _
    Or currentROPDOSISRV = Trim(xYROPINF0.ROPINFGUSR) _
    Or currentROPDOSISRV = Trim(xYROPINF0.ROPINFGSRV) Then
       blnROPINFGPRV = True
    Else
       blnROPINFGPRV = False
    End If
Else
    blnROPINFGPRV = True
End If

wUsr_Aut = ""

Dossier_Aut = False_Aut
Processus_Aut = False_Aut
Action_Aut = False_Aut
Memo_Aut = False_Aut
If DROPI_Aut.Xspécial Then
    Dossier_Aut.Saisir = True
    Dossier_Aut.Valider = True
    Processus_Aut = Dossier_Aut
    Action_Aut = Dossier_Aut
    Memo_Aut = Dossier_Aut
Else

    If cmdUpdate_Init_00_Hab(Trim(oldYROPDOS0.ROPDOSGUSR)) Or usrName_UCase = Trim(oldYROPDOS0.ROPDOSGUSR) Then
        Dossier_Aut.Saisir = True
        Dossier_Aut.Valider = True
        Processus_Aut = Dossier_Aut
        Action_Aut = Dossier_Aut
        Memo_Aut = Dossier_Aut
    Else
        If cmdUpdate_Init_00_Hab(Trim(oldYROPDOS0.ROPDOSIUSR)) Or usrName_UCase = Trim(oldYROPDOS0.ROPDOSIUSR) Then
            Dossier_Aut.Saisir = True
            Dossier_Aut.Valider = True
            Processus_Aut = Dossier_Aut
            Action_Aut = Dossier_Aut
            Memo_Aut = Dossier_Aut
        Else
            If cmdUpdate_Init_00_Hab(Trim(arrYROPINF0(Processus_Index).ROPINFGUSR)) Then
                Processus_Aut.Saisir = True
                Processus_Aut.Valider = True
                Action_Aut = Processus_Aut
                Memo_Aut = Processus_Aut
            Else
                If cmdUpdate_Init_00_Hab(Trim(arrYROPINF0(Action_Index).ROPINFGUSR)) Then
                    Action_Aut.Saisir = True
                    Action_Aut.Valider = True
                    Memo_Aut = Action_Aut
                Else
                    If cmdUpdate_Init_00_Hab(Trim(oldYROPINF0.ROPINFGUSR)) Then
                        Memo_Aut.Saisir = True
                        Memo_Aut.Valider = True
                    End If
                End If
            End If
        End If
    End If
End If
'________________________________________________________________
If oldYROPDOS0.ROPDOSSTA <> " " Then
    Processus_Aut = False_Aut
    Action_Aut = False_Aut
    Memo_Aut = False_Aut
Else
    If arrYROPINF0(Processus_Index).ROPINFSTA <> " " Or arrYROPINF0(Processus_Index).ROPINFSTAD <> " " Then
        Action_Aut = False_Aut
        Memo_Aut = False_Aut
    Else
        If arrYROPINF0(Action_Index).ROPINFSTA <> " " Or arrYROPINF0(Action_Index).ROPINFSTAD <> " " Then
            Memo_Aut = False_Aut
        End If
    End If
End If
'________________________________________________________________
blnDossierSAnsModèle = True
blnProcessusAvecActionFinale = False
For K = 1 To arrYROPINF0_Nb
    If arrYROPINF0(K).ROPINFIDP > 1 Then blnDossierSAnsModèle = False ': Exit For
    If arrYROPINF0(K).ROPINFIDT > 0 Then blnDossierSAnsModèle = False ': Exit For
    If arrYROPINF0(K).ROPINFIDP = xYROPINF0.ROPINFIDP And arrYROPINF0(K).ROPINFGNAT = "F" Then blnProcessusAvecActionFinale = True
Next K

'________________________________________________________________ habilitations
arrSelect_Update_Nb = 0 '1: arrSelect_Update(arrSelect_Update_Nb) = "00 - Afficher"
Select Case cmdUpdate_Init_K
'_________________________________________________________________________________________________________
    Case "D"
        If blnROPDOSQUAL Then cmdUpdate_Dossier.AddItem "14Q- Modifier la qualification": cmdUpdate_Dossier.AddItem "14 - Modifier ce dossier"

            If oldYROPDOS0.ROPDOSSTA = " " Then
              
                If Dossier_Aut.Saisir Then
                    cmdUpdate_Dossier.AddItem "03 - Ajouter un processus"
                End If
                If Dossier_Aut.Valider Then
                    If blnDossierSAnsModèle Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "74 - choisir un modèle de gestion"
                    cmdUpdate_Dossier.AddItem "14 - Modifier ce dossier"
                    cmdUpdate_Dossier.AddItem "24 - Clôturer ce dossier"
                    cmdUpdate_Dossier.AddItem "34 - Annuler ce dossier"
                    cmdUpdate_Dossier.AddItem "64 - Report d'échéance du dossier + actions"
                    If blnDossierSAnsModèle Then cmdUpdate_Dossier.AddItem "74 - choisir un modèle de gestion"
                 End If
            Else
                If Dossier_Aut.Valider Then cmdUpdate_Dossier.AddItem "54 - Réactiver ce dossier"
            End If
            
            If Dossier_Aut.Xspécial Then cmdUpdate_Dossier.AddItem "44 - Effacer ce dossier"
    
    
'_________________________________________________________________________________________________________
    Case "P"
             If xYROPINF0.ROPINFSTA = " " And xYROPINF0.ROPINFSTAD = " " Then
               If blnROPINFGPRV Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "13 - Modifier ce processus"
               If Processus_Aut.Saisir Then
                    If Not blnProcessusAvecActionFinale Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "02F- Ajouter une action pour la fermeture du processus"
                End If
                If Processus_Aut.Valider Then
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "23 - Clôturer ce processus"
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "33 - Annuler ce processus"
                End If
             Else
                If Processus_Aut.Valider Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "53 - Réactiver ce processus"
            End If
           If Processus_Aut.Xspécial Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "43 - Effacer ce processus"
'_________________________________________________________________________________________________________
  Case "A", "F"
            If xYROPINF0.ROPINFSTA = " " And xYROPINF0.ROPINFSTAD = " " Then
                If blnROPINFGPRV Then
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "12 - Modifier cette action": arrSelect_Update_K = arrSelect_Update_Nb
                    blnYROPINF0_12X_Aut = True
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "12X- Réaffecter(Resp,Ech) cette action"
                End If
                If Action_Aut.Saisir Then
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "02I- Insérer une action avant cette action"
                End If
                If Action_Aut.Valider Then
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "22 - Clôturer cette action"
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "22 + Clôturer cette action + ajouter une autre action"
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "32 - Annuler cette action"
                End If
            Else
                If Action_Aut.Valider Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "52 - Réactiver cette action"
            End If
            If Action_Aut.Xspécial Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "42 - Effacer cette action"
'_________________________________________________________________________________________________________
    Case "J"
            If xYROPINF0.ROPINFSTA = " " And xYROPINF0.ROPINFSTAD = " " Then
                arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "31 - Annuler cette  pièce jointe"
            End If
            
            If Memo_Aut.Xspécial Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "41 - Effacer cette pièce jointe"
'_________________________________________________________________________________________________________
    Case Else
            If xYROPINF0.ROPINFSTA = " " And xYROPINF0.ROPINFSTAD = " " Then
                If blnROPINFGPRV Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "11 - Modifier cette note": arrSelect_Update_K = arrSelect_Update_Nb
                arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "31 - Annuler cette note"
            Else
                If Memo_Aut.Valider Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "51 - Réactiver cette note"
            End If
            
            If Memo_Aut.Xspécial Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "41 - Effacer cette note"
End Select


If Not blnROPINFIDTL_Ok Then
    If Not blnYROPINF0_12X_Aut Then
        cmdUpdate.Enabled = False
    Else
        arrSelect_Update_Nb = 1: arrSelect_Update(arrSelect_Update_Nb) = "12X- Réaffecter(Resp,Ech) cette action"
        arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "02I- Insérer une action avant cette action"
    End If
End If


cmdUpdate_01.Visible = False
cmdUpdate_02.Visible = False
cmdUpdate_05.Visible = False
Select Case oldYROPDOS0.ROPDOSSTA
    Case "A", "F"
    Case Else: cmdUpdate_01.Visible = Action_Aut.Saisir
               cmdUpdate_02.Visible = Action_Aut.Saisir
               cmdUpdate_05.Visible = Action_Aut.Saisir

End Select
'fraDossier_B.Enabled = False
fraUpdate_B.Enabled = False
cmdDossier_Ok.Visible = False: fraDossier_cmd.Visible = False 'True
cmdDossier_Ok_Close.Visible = False
cmdUpdate_Display (arrSelect_Update(1))



End Sub

Public Sub fraDossier_STAK_Processus(kProcessus As Long)

If kProcessus > 0 Then
    Select Case arrYROPINF0(kProcessus).ROPINFSTA
        Case "F": arrYROPINF0(kProcessus).ROPINFSTAK = "V"
        Case "A": arrYROPINF0(kProcessus).ROPINFSTAK = "A"
        Case " ":
            blnProcessus_EnCours = True
            
            If arrYROPINF0(kProcessus).ROPINFGECH > DSys Then
                If blnAction_EnAlerte Then
                    arrYROPINF0(kProcessus).ROPINFSTAK = "!": blnProcessus_EnAlerte = True
                Else
                    arrYROPINF0(kProcessus).ROPINFSTAK = "B"
               End If
                
            Else
                blnProcessus_EnAlerte = True
                If arrYROPINF0(kProcessus).ROPINFGECH = DSys Then
                    If blnAction_EnAlerte Then
                        arrYROPINF0(kProcessus).ROPINFSTAK = "!"
                    Else
                        arrYROPINF0(kProcessus).ROPINFSTAK = "O"
                    End If
                Else
                    If blnAction_Valide Then
                        If blnAction_EnAlerte Then
                            arrYROPINF0(kProcessus).ROPINFSTAK = "!"
                        Else
                            arrYROPINF0(kProcessus).ROPINFSTAK = "R"
                        End If
                    Else
                        arrYROPINF0(kProcessus).ROPINFSTAK = "M"
                    End If
                End If
            End If
        Case Else: arrYROPINF0(kProcessus).ROPINFSTAK = " "
    End Select
End If
blnAction_EnCours = False: blnAction_EnAlerte = False: blnAction_Valide = False

End Sub

Private Sub txtUpdate_ROPDOSGPRV_GotFocus()
txt_GotFocus txtUpdate_ROPDOSGPRV

End Sub

Private Sub txtUpdate_ROPDOSGPRV_LostFocus()
txt_LostFocus txtUpdate_ROPDOSGPRV

End Sub


Private Sub txtUpdate_ROPDOSGUSR_Click()
Dim X As String
If blnControl Then
    chkUpdate_ROPINFMAIL_D.Caption = fraDossier_Display_USR(txtUpdate_ROPDOSGUSR)
    Call cmdDossier_Ok_GSRV(txtUpdate_ROPDOSGUSR.Text, X)
    If Mid$(X, 1, 2) = "_S" Then libUpdate_ROPDOSGSRV = arrROPDOSISRV_Lib(Val(Mid$(X, 3, 2)))

End If
End Sub

Private Sub txtUpdate_ROPDOSGUSR_GotFocus()
txt_GotFocus txtUpdate_ROPDOSGUSR

End Sub


Private Sub txtUpdate_ROPDOSGUSR_LostFocus()
txt_LostFocus txtUpdate_ROPDOSGUSR

End Sub


Private Sub txtUpdate_ROPDOSIAMJ_GotFocus()
lblUpdate_ROPDOSIAMJ.BackColor = focusUsr.BackColor

End Sub


Private Sub txtUpdate_ROPDOSIAMJ_LostFocus()
lblUpdate_ROPDOSIAMJ.BackColor = fraDossier_C.BackColor '&HC0F0FF

End Sub


Private Sub txtUpdate_ROPDOSIREF_GotFocus()
txt_GotFocus txtUpdate_ROPDOSIREF

End Sub


Private Sub txtUpdate_ROPDOSIREF_LostFocus()
txt_LostFocus txtUpdate_ROPDOSIREF

End Sub


Private Sub txtUpdate_ROPDOSIUSR_Click()
Dim X As String
If blnControl Then
    chkUpdate_ROPINFMAIL_I.Caption = fraDossier_Display_USR(txtUpdate_ROPDOSIUSR)
    Call cmdDossier_Ok_GSRV(txtUpdate_ROPDOSIUSR.Text, X)
    If Mid$(X, 1, 2) = "_S" Then libUpdate_ROPDOSISRV = arrROPDOSISRV_Lib(Val(Mid$(X, 3, 2)))
End If


End Sub

Private Sub txtUpdate_ROPDOSIUSR_GotFocus()
txt_GotFocus txtUpdate_ROPDOSIUSR

End Sub


Private Sub txtUpdate_ROPDOSIUSR_LostFocus()
txt_LostFocus txtUpdate_ROPDOSIUSR

End Sub


Private Sub txtUpdate_ROPDOSXAPP_GotFocus()
txt_GotFocus txtUpdate_ROPDOSXAPP

End Sub


Private Sub txtUpdate_ROPDOSXAPP_LostFocus()
txt_LostFocus txtUpdate_ROPDOSXAPP

End Sub


Private Sub txtUpdate_ROPDOSXDOM_Click()
If txtUpdate_ROPDOSXDOM.Enabled Then sqlYBIATAB0_cboID_K1 "ROPDOSXAPP", Mid$(txtUpdate_ROPDOSXDOM, 1, 12), txtUpdate_ROPDOSXAPP

End Sub

Private Sub txtUpdate_ROPDOSXDOM_GotFocus()
txt_GotFocus txtUpdate_ROPDOSXDOM

End Sub


Private Sub txtUpdate_ROPDOSXDOM_LostFocus()
txt_LostFocus txtUpdate_ROPDOSXDOM

End Sub


Private Sub txtUpdate_ROPDOSXID_GotFocus()
txt_GotFocus txtUpdate_ROPDOSXID
End Sub

Private Sub txtUpdate_ROPDOSXID_LostFocus()
txt_LostFocus txtUpdate_ROPDOSXID

End Sub

Private Sub txtUpdate_ROPINFGECH_GotFocus()
txtUpdate_ROPINFGECH_Old = txtUpdate_ROPINFGECH

lblUpdate_ROPINFGECH.BackColor = fraUpdate_B.BackColor
End Sub

Private Sub txtUpdate_ROPINFGECH_LostFocus()
If cmdSelect_SQL_K = "2" _
And txtUpdate_ROPINFGECH_Old <> txtUpdate_ROPINFGECH Then blntxtUpdate_ROPINFGECH_Change = True
lblUpdate_ROPINFGECH.BackColor = fraUpdate_B.BackColor
End Sub

Private Sub txtUpdate_ROPINFGTXT_Click()
'If fraDossier_B.Enabled Then
If Not txtUpdate_ROPINFGTXT.Locked Then
    libDossier_ROPINFGTXT.Visible = True
End If
End Sub


Private Sub txtUpdate_ROPINFGTXT_GotFocus()
txt_GotFocus txtUpdate_ROPINFGTXT

End Sub


Private Sub txtUpdate_ROPINFGTXT_LostFocus()
txt_LostFocus txtUpdate_ROPINFGTXT
Select Case cmdSelect_SQL_K
    Case "0", "2", "2M": '' chkSelect_Update_B.Value = "1"
'    Case Else: libDossier_ROPINFGTXT.Visible = False
End Select
End Sub


Private Sub txtUpdate_ROPINFGUO_GotFocus()
txt_GotFocus txtUpdate_ROPINFGUO

End Sub

Private Sub txtUpdate_ROPINFGUO_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtUpdate_ROPINFGUO)
End Sub



Public Sub cmdDossier_Ok_04_ISRV()
Dim X As String, K As Integer
''''If IsNull(sqlYBIATAB0_Read("ROPDOSGUSR", newYROPDOS0.ROPDOSIUSR, "", X)) Then

X = "select SSIDOMUNIT from " & paramIBM_Library_SABSPE & ".YSSIDOM0 where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN'" _
     & " and SSIDOMUIDX = '" & newYROPDOS0.ROPDOSIUSR & "'"
Set rsSab = cnsab.Execute(X)

If Not rsSab.EOF Then
    newYROPDOS0.ROPDOSISRV = "_" & rsSab("SSIDOMUNIT")
Else
    newYROPDOS0.ROPDOSISRV = ""
End If
Call cmdDossier_Ok_GSRV(newYROPDOS0.ROPDOSIUSR, newYROPDOS0.ROPDOSISRV)
End Sub

Public Sub cmdDossier_Ok_GSRV(lUsr As String, lSRV As String)
Dim X As String, K As Integer
If Mid$(lUsr, 1, 2) = "_S" Then
    lSRV = lUsr
Else
    X = "select SSIDOMUNIT from " & paramIBM_Library_SABSPE & ".YSSIDOM0 where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN'" _
         & " and SSIDOMUIDX = '" & lUsr & "'"
    Set rsSab = cnsab.Execute(X)
    
    If Not rsSab.EOF Then
        lSRV = "_" & rsSab("SSIDOMUNIT")
    Else
        lSRV = ""
    End If
End If

End Sub

Public Sub cmdPrint0_Dossier()
Dim mROPINFIDP As Long
Dim V, K As Long, X As String


prtDROPI_Dossier oldYROPDOS0
mROPINFIDP = 1
For K = 1 To arrYROPINF0_Nb
    prtDROPI_Détail_1 arrYROPINF0(K)
Next K

XPrt.DrawWidth = 8
XPrt.CurrentY = XPrt.CurrentY + 100
XPrt.Line (prtMinX, XPrt.CurrentY + prtlineHeight)-(prtMaxX, XPrt.CurrentY + prtlineHeight), prtLineColor
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight + 20


End Sub

Public Sub cmdPrint0_Echéancier()
Dim mROPINFIDP As Long
Dim V, K As Long, X As String
Dim mYROPPRT0 As typeYROPPRT0, xYROPPRT0 As typeYROPPRT0
Dim blnOpen As Boolean, blnprtDROPI_Détail_2P As Boolean
Dim xSQL As String

mYROPPRT0.ROPPRTDEST = ""
blnOpen = False
'MsgBox "cmdPrint0_Echéancier 10 / " & lstW.ListCount
For K = 0 To lstW.ListCount - 1
    lstW.ListIndex = K
    xYROPPRT0.ROPPRTDEST = Mid$(lstW.Text, 1, 12)
    xYROPPRT0.ROPPRTID = Val(Mid$(lstW.Text, 14, 9))
    xYROPPRT0.ROPPRTIDP = Val(Mid$(lstW.Text, 24, 9))
    xYROPPRT0.ROPPRTIDT = Val(Mid$(lstW.Text, 34, 9))
    xYROPPRT0.ROPPRTarrI = Val(Mid$(lstW.Text, 44, 9))
    xYROPPRT0.ROPPRTarrD = Val(Mid$(lstW.Text, 54, 9))
    xYROPPRT0.ROPPRTGECH = Mid$(lstW.Text, 64, 8)
    xYROPPRT0.ROPPRTGUSR = Mid$(lstW.Text, 73, 12)
    
    blnprtDROPI_Détail_2P = False
    If mYROPPRT0.ROPPRTDEST <> xYROPPRT0.ROPPRTDEST Then
        If mYROPPRT0.ROPPRTDEST <> "" Then cmdPrint0_Echéancier_Destinataire
        blnOpen = True
        blnprtDROPI_Détail_2P = True
        prtDROPI_Open 2, "Risques Opérationnels : Echéancier " & xYROPPRT0.ROPPRTDEST & " au " & DSys
    End If
    
    If mYROPPRT0.ROPPRTID <> xYROPPRT0.ROPPRTID _
    Or mYROPPRT0.ROPPRTIDP <> xYROPPRT0.ROPPRTIDP Then
        blnprtDROPI_Détail_2P = True
        xSQL = "select *  from " & paramIBM_Library_SABSPE & ".YROPINF0 " _
         & " where ROPINFID = " & xYROPPRT0.ROPPRTID _
         & " and ROPINFIDP = " & xYROPPRT0.ROPPRTIDP _
         & " and ROPINFIDT = " & 0 _
         & " and ROPINFIDT2 = " & 1

        Set rsSab = cnsab.Execute(xSQL)
        If rsSab.EOF Then
            Call rsYROPINF0_Init(xYROPINF0)
        Else
            V = rsYROPINF0_GetBuffer(rsSab, xYROPINF0)
            If Not IsNull(V) Then Call rsYROPINF0_Init(xYROPINF0)
        End If
    End If
      
   ' If blnprtDROPI_Détail_2P Then prtDROPI_Détail_2P arrYROPDOS0(xYROPPRT0.ROPPRTarrD), xYROPINF0
   ' Call prtDROPI_Détail_2A(arrYROPINF0(xYROPPRT0.ROPPRTarrI), blnprtDROPI_Détail_2P)
   
    If blnprtDROPI_Détail_2P Then
        prtDROPI_Dossier arrYROPDOS0(xYROPPRT0.ROPPRTarrD)
        Call prtDROPI_Détail_2(xYROPINF0)
    End If
    Call prtDROPI_Détail_2(arrYROPINF0(xYROPPRT0.ROPPRTarrI))
   
    mYROPPRT0 = xYROPPRT0

Next K


If blnOpen Then cmdPrint0_Echéancier_Destinataire


End Sub

Public Sub cmdSendMail_Echéancier()
Dim mROPINFIDP As Long
Dim V, K As Long, X As String
Dim mYROPPRT0 As typeYROPPRT0, xYROPPRT0 As typeYROPPRT0
Dim blnOpen As Boolean, blnprtDROPI_Détail_2P As Boolean
Dim xSQL As String

mYROPPRT0.ROPPRTDEST = ""
blnOpen = False

For K = 0 To lstW.ListCount - 1
    lstW.ListIndex = K
    xYROPPRT0.ROPPRTDEST = Mid$(lstW.Text, 1, 12)
    xYROPPRT0.ROPPRTID = Val(Mid$(lstW.Text, 14, 9))
    xYROPPRT0.ROPPRTIDP = Val(Mid$(lstW.Text, 24, 9))
    xYROPPRT0.ROPPRTIDT = Val(Mid$(lstW.Text, 34, 9))
    xYROPPRT0.ROPPRTarrI = Val(Mid$(lstW.Text, 44, 9))
    xYROPPRT0.ROPPRTarrD = Val(Mid$(lstW.Text, 54, 9))
    xYROPPRT0.ROPPRTGECH = Mid$(lstW.Text, 64, 8)
    xYROPPRT0.ROPPRTGUSR = Mid$(lstW.Text, 73, 12)
    
    blnprtDROPI_Détail_2P = False
    If mYROPPRT0.ROPPRTDEST <> xYROPPRT0.ROPPRTDEST Then
        If mYROPPRT0.ROPPRTDEST <> "" Then cmdSendMail_Echéancier_Destinataire
        blnOpen = True
        blnprtDROPI_Détail_2P = True
        Call cmdSendMail_Echéancier_Open(xYROPPRT0.ROPPRTDEST)
    End If
    
    If mYROPPRT0.ROPPRTID <> xYROPPRT0.ROPPRTID _
    Or mYROPPRT0.ROPPRTIDP <> xYROPPRT0.ROPPRTIDP Then
        blnprtDROPI_Détail_2P = True
        xSQL = "select *  from " & paramIBM_Library_SABSPE & ".YROPINF0 " _
         & " where ROPINFID = " & xYROPPRT0.ROPPRTID _
         & " and ROPINFIDP = " & xYROPPRT0.ROPPRTIDP _
         & " and ROPINFIDT = " & 0 _
         & " and ROPINFIDT2 = " & 1
        Set rsSab = cnsab.Execute(xSQL)
        If rsSab.EOF Then
            Call rsYROPINF0_Init(xYROPINF0)
        Else
            V = rsYROPINF0_GetBuffer(rsSab, xYROPINF0)
            If Not IsNull(V) Then Call rsYROPINF0_Init(xYROPINF0)
        End If
    End If
      
   
    If blnprtDROPI_Détail_2P Then
        xYROPDOS0 = arrYROPDOS0(xYROPPRT0.ROPPRTarrD)
        Call cmdSendMail_Echéancier_Dossier
        Call cmdSendMail_Echéancier_Détail
    End If
    xYROPINF0 = arrYROPINF0(xYROPPRT0.ROPPRTarrI)
    Call cmdSendMail_Echéancier_Détail
   
    mYROPPRT0 = xYROPPRT0

Next K


If blnOpen Then cmdSendMail_Echéancier_Destinataire


End Sub

Public Sub fraDétail_lbl(lROPINFGNAT As String, lROPINFIDP As Long)
Select Case lROPINFGNAT
    Case "P":
         lblUpdate_ROPINFGECH.Caption = "Echéance"
             lblUpdate_ROPINFGUSR.Caption = "Responsable"
     Case "A"
        lblUpdate_ROPINFGUSR.Caption = "Responsable"
        lblUpdate_ROPINFGECH.Caption = "Echéance"
      Case "N"
        lblUpdate_ROPINFGUSR.Caption = "Auteur"
        lblUpdate_ROPINFGECH.Caption = "Date"
       Case "J"
        lblUpdate_ROPINFGUSR.Caption = "Pièce Jointe"
        lblUpdate_ROPINFGECH.Caption = "Date"
         
    Case Else
        lblUpdate_ROPINFGUSR.Caption = "??"
        lblUpdate_ROPINFGECH.Caption = "date"
End Select

End Sub

Public Sub fraParam_Reset()
lstParam_ROPDOSXAPP.Enabled = False
fraParam_Update.Visible = False
rsYBIATAB0_Init oldParam
txtParam_BIATABK1 = ""
txtParam_BIATABK2 = ""
txtParam_BIATABTXT = ""
libROPDOSMAIL.Visible = False
txtROPDOSMAIL.Visible = False
blnROPDOSMAIL = False
End Sub
Public Sub fraAut_Reset()

On Error Resume Next

fraAut_Update.Visible = False
rsYBIATAB0_Init oldParam
'Call lstZMNURUT0_Load_Actif_Production(lstAut_Usr)
Call YSSIUSR0_Actif_Load(lstAut_Usr)
End Sub

Public Sub fraAut_Control()
Dim X As String
Dim wIndex As Long
On Error Resume Next
X = " "
If optAut_ROPDOSISRV_H Then X = "H"
If optAut_ROPDOSISRV_R Then X = "R"
If optAut_ROPDOSISRV_D Then X = "D"
If optAut_ROPDOSISRV_C Then X = "C"
If optAut_ROPDOSISRV_X Then X = "X"
If optAut_ROPDOSISRV_I Then X = "I"

Mid$(newAut.BIATABTXT, arrROPDOSISRV_K, 1) = X
'lstAut_ROPDOSGUSR_Display (wIndex)

End Sub

Public Sub fraParam_Display()
fraParam_Update.Visible = True

txtParam_BIATABK1 = Trim(oldParam.BIATABK1): txtParam_BIATABK1.Enabled = False

If Trim(oldParam.BIATABID) = "ROPDOSISRV" Then
    txtParam_BIATABK2 = Trim(Mid$(oldParam.BIATABTXT, 1, 12)): txtParam_BIATABK2.Enabled = True
    txtParam_BIATABTXT = Trim(Mid$(oldParam.BIATABTXT, 13, 64)): txtParam_BIATABTXT.Enabled = True
Else
    txtParam_BIATABK2 = Trim(oldParam.BIATABK2): txtParam_BIATABK2.Enabled = False
    txtParam_BIATABTXT = Trim(oldParam.BIATABTXT): txtParam_BIATABTXT.Enabled = False
    If Trim(oldParam.BIATABID) = "ROPDOSQUAL" And Not IsNumeric(Mid$(oldParam.BIATABTXT, 1, 1)) Then
        txtParam_BIATABTXT = "9." & Trim(oldParam.BIATABTXT)
    End If
End If
End Sub
Public Function fraParam_Update_Control()
Dim blnUpdate_Control As Boolean
Dim X As String, wMsg As String
Dim bln
blnUpdate_Control = True
newParam = oldParam
wMsg = ""

X = Trim(txtParam_BIATABK1)
newParam.BIATABK1 = X

If Trim(newParam.BIATABID) = "ROPDOSISRV" Then
    newParam.BIATABK1 = UCase$(X)
    If Len(X) <> 3 Then blnUpdate_Control = False
    If Not IsNumeric(Mid$(X, 2, 2)) Then blnUpdate_Control = False
    If blnUpdate_Control = False Then
        wMsg = wMsg & "- le code doit être au format S** (** numérique)" & vbCrLf
    End If

    X = Trim(txtParam_BIATABK2)
    If X = "" Then
        blnUpdate_Control = False
        wMsg = wMsg & "- préciser le libellé réduit" & vbCrLf
    Else
        Mid$(newParam.BIATABTXT, 1, 12) = Space$(12)
        Mid$(newParam.BIATABTXT, 1, 12) = X
    End If

    X = Trim(txtParam_BIATABTXT)
    If X = "" Then
        blnUpdate_Control = False
        wMsg = wMsg & "- préciser le libellé" & vbCrLf
    Else
        Mid$(newParam.BIATABTXT, 13, 64) = Space$(52)
        Mid$(newParam.BIATABTXT, 13, 64) = X
    End If
    GoTo fraParam_Update_Control_End
End If


Select Case Trim(newParam.BIATABID)
    Case "ROPDOSQUAL":
                newParam.BIATABK1 = UCase$(X)
                If Len(X) <> 3 Then
                    blnUpdate_Control = False
                    wMsg = wMsg & "- la référence RO doit avoir 3 caractères" & vbCrLf
                End If
    Case "ROPDOSQUALB2":
                newParam.BIATABK1 = UCase$(X)
                If Len(X) <> 1 Or Not IsNumeric(X) Then
                    blnUpdate_Control = False
                    wMsg = wMsg & "- la référence Bâle II = (1-9)" & vbCrLf
                End If
End Select

    X = Trim(txtParam_BIATABK2)
    newParam.BIATABK2 = X
    
    X = Trim(txtParam_BIATABTXT)
    newParam.BIATABTXT = X
    If X = "" Then
        blnUpdate_Control = False
        wMsg = wMsg & "- préciser le libellé" & vbCrLf
    End If
    
    Select Case Trim(newParam.BIATABID)
        Case "ROPDOSQUAL":
                    If Not IsNumeric(Mid$(newParam.BIATABTXT, 1, 1)) Then
                        blnUpdate_Control = False
                        wMsg = wMsg & "- la liaison à la référence Bâle II  = (1-9)" & vbCrLf
                    End If
                    If Mid$(newParam.BIATABTXT, 2, 1) <> "." Then
                        blnUpdate_Control = False
                        wMsg = wMsg & "- le deuxième caractère du libellé doit être un point ." & vbCrLf
                    End If
    End Select
    

fraParam_Update_Control_End:

If blnUpdate_Control Then
    fraParam_Update_Control = Null
Else
    fraParam_Update_Control = "<Fin du contrôle des données "
    Call MsgBox(wMsg, vbCritical, "DROPI : paramétrage")

End If
End Function


Public Sub Form_Init_RODOSGUSR()
Dim X As String, xK1 As String, xUsr As String, xHabilitation As String
Dim K As Integer

For K = 0 To 100
    arrROPDOSISRV_Mail(K) = ""
Next K

blnROPDOSQUAL = False: blnExportation_xlsx = False

txtUpdate_ROPDOSGUSR.Clear
txtUpdate_ROPDOSGUSR.AddItem "?"
txtUpdate_ROPDOSIUSR.Clear
txtUpdate_ROPDOSIUSR.AddItem "?"

txtUpdate_ROPINFGUSR.Clear
txtUpdate_ROPINFGUSR.AddItem "?"
lstUpdate_ROPINFMAIL.Clear
lstUpdate_ROPINFMAIL_CC.Clear

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0, " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
  & " where BIATABID = 'ROPDOSGUSR'" _
  & " and SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN' and SSIDOMUIDX = BIATABK1 order by BIATABK1"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    xK1 = rsSab("BIATABK1")
    xUsr = Trim(xK1)
    X = rsSab("BIATABTXT")
    If Mid$(rsSab("SSIDOMUNIT"), 1, 1) = "S" Then
        K = Val(Mid$(rsSab("SSIDOMUNIT"), 2, 2))
        xHabilitation = Mid$(X, 28 + K, 1)
        If xHabilitation = "R" Or xHabilitation = "D" Then arrROPDOSISRV_Mail(K) = arrROPDOSISRV_Mail(K) & xK1
        If Mid$(X, 1, 1) = "D" Then txtUpdate_ROPDOSGUSR.AddItem xUsr
        If Mid$(X, 2, 1) = "A" Then txtUpdate_ROPINFGUSR.AddItem xUsr: txtUpdate_ROPDOSIUSR.AddItem xUsr
        If xUsr = usrName_UCase Then
            currentROPDOSISRV = "_" & rsSab("SSIDOMUNIT")
            currentROPDOSISRV_Nom = Trim(fraDossier_Display_USR(currentROPDOSISRV))
            currentROPDOSISRV_Hab = Mid$(X, 29, 76)
            currentROPDOSISRV_Rôle = xHabilitation
            If Mid$(X, 3, 1) = "P" Then
                fraParam.Visible = True
                mnuExport_Param.Visible = True
            Else
                fraParam.Visible = False
            End If
            If Mid$(X, 4, 1) = "H" Then
                fraAut.Visible = True
            Else
                fraAut.Visible = False
            End If
            If Mid$(X, 5, 1) = "Q" Then
                blnROPDOSQUAL = True
            Else
                blnROPDOSQUAL = False
            End If
             If Mid$(X, 6, 1) = "E" Then
                blnExportation_xlsx = True
            Else
                blnExportation_xlsx = False
            End If
           
        End If
        lstUpdate_ROPINFMAIL.AddItem xUsr 'Mid$(xUsr, 1, 10)
        lstUpdate_ROPINFMAIL_CC.AddItem xUsr 'Mid$(xUsr, 1, 10)
    End If
    rsSab.MoveNext
Loop

ReDim arrROPINFMAIL(lstUpdate_ROPINFMAIL.ListCount)
For K = 0 To lstUpdate_ROPINFMAIL.ListCount
    lstUpdate_ROPINFMAIL.ListIndex = K
    arrROPINFMAIL(K) = lstUpdate_ROPINFMAIL.Text
Next K
'________________________________________________________________________________
X = "select *from " & paramIBM_Library_SABSPE & ".YSSIUSR0 where SSIUSRNAT= 'S' order by SSIUSRUNIT"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    xUsr = "_" & Format(rsSab("SSIUSRUNIT"), "          ") & "_" & Trim(rsSab("SSIUSRPRFX"))
    txtUpdate_ROPDOSGUSR.AddItem xUsr
    txtUpdate_ROPDOSIUSR.AddItem xUsr
    
    txtUpdate_ROPINFGUSR.AddItem xUsr
    rsSab.MoveNext
Loop


End Sub

Public Sub Form_Init_RODOSISRV()
Dim X As String, K As Integer
Dim xK1 As String
For K = 1 To 100
    arrROPDOSISRV_K1(K) = ""
    arrROPDOSISRV_Code(K) = "?" & K
    arrROPDOSISRV_Lib(K) = "?" & K
    arrROPDOSISRV_ListIndex(K) = -1
Next K

lstAut_ROPDOSISRV.Clear
lstParam_ROPDOSISRV.Clear
X = "select *from " & paramIBM_Library_SABSPE & ".YSSIUSR0 where SSIUSRNAT= 'S' order by SSIUSRUNIT"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    xK1 = Trim(rsSab("SSIUSRUNIT"))
    K = Val(Mid$(xK1, 2, 2))
    arrROPDOSISRV_K1(K) = xK1
    'X = rsSab("BIATABTXT")
    arrROPDOSISRV_Code(K) = Trim(rsSab("SSIUSRPRFX"))
    arrROPDOSISRV_Lib(K) = Trim(rsSab("SSIUSRUIDX"))
    lstAut_ROPDOSISRV.AddItem xK1 & " " & arrROPDOSISRV_Code(K) & " " & arrROPDOSISRV_Lib(K)
    lstParam_ROPDOSISRV.AddItem xK1 & " - " & arrROPDOSISRV_Code(K) & " - " & arrROPDOSISRV_Lib(K)
    arrROPDOSISRV_ListIndex(K) = lstAut_ROPDOSISRV.ListCount
    rsSab.MoveNext
Loop

End Sub

Public Sub cmdUpdate_Display(lItem As String)
Dim K As Integer
cmdUpdate.Clear
For K = 1 To arrSelect_Update_Nb
    cmdUpdate.AddItem arrSelect_Update(K)
    Select Case Mid$(arrSelect_Update(K), 1, 4)
        Case "12 -", "11 -", "13 -": cmdUpdate_12.Visible = True: cmdUpdate_12.Caption = arrSelect_Update(K)
        Case "22 -", "21 -", "23 -": cmdUpdate_22.Visible = True: cmdUpdate_22.Caption = arrSelect_Update(K)
        Case "32 -", "31 -", "33 -": cmdUpdate_32.Visible = True: cmdUpdate_32.Caption = arrSelect_Update(K)
    End Select
    'If lItem = arrSelect_Update(K) Then
    '    cmdUpdate.AddItem "=>" & arrSelect_Update(K)
    'Else
    '    cmdUpdate.AddItem "  " & arrSelect_Update(K)
    'End If
    
Next K
End Sub

Public Sub cmdSelect_SQL_2M()
X = InputBox("Indiquez le numéro de dossier (< 1000)")
DossierModèle_ROPDOSID = Val(X)
If DossierModèle_ROPDOSID > 0 And DossierModèle_ROPDOSID < 1000 Then
    'cmdSelect_SQL_K = 2
    blnDossierModèle = True
    rsYROPDOS0_Init xYROPDOS0
    xYROPDOS0.ROPDOSXDOM = "Modèle"
    xYROPDOS0.ROPDOSGPRV = "V"
    oldYROPDOS0 = xYROPDOS0
    rsYROPINF0_Init xYROPINF0
    xYROPINF0.ROPINFGNAT = "P"
    xYROPINF0.ROPINFMAIL = "D"
    oldYROPINF0 = xYROPINF0
    arrYROPINF0_Nb = 0

    cmdSelect_SQL_2
Else
    MsgBox "N° de dossier invalide"
End If
End Sub

Public Sub cmdSelect_SQL_1_YROPDOS0()
Dim K As Integer, blnOk As Boolean
Dim xSQL As String
Dim Xdisplay As String
Dim X As String
Dim HeightOfLine As Long, LinesOfText As Long

currentAction = "cmdSelect_SQL_1_YROPDOS0"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fraUpdate_B.Visible = False

For K = 1 To selYROPDOS0_Nb
    xYROPDOS0 = selYROPDOS0(K)
    blnOk = False
    If xYROPDOS0.ROPDOSGPRV = "U" Then
        If Trim(xYROPDOS0.ROPDOSGUSR) = usrName_UCase Then blnOk = True
    Else
        If cmdSelect_SQL_K = "1X" Then
            blnOk = True
        Else
            Select Case xYROPDOS0.ROPDOSGPRV
                Case "W": blnOk = True
                Case Else:
                    If cmdSelect_SQL_1_Hab(Trim(xYROPDOS0.ROPDOSGSRV)) Then
                        blnOk = True
                    Else
                        If cmdSelect_SQL_1_Hab(Trim(xYROPDOS0.ROPDOSISRV)) Then blnOk = True
                    End If
            End Select
        End If
    End If
    
    If blnOk Then
        xSQL = "select ROPINFGTXT from " & paramIBM_Library_SABSPE & ".YROPINF0" _
            & " where ROPINFID = " & xYROPDOS0.ROPDOSID _
            & " and ROPINFIDP = 1 and ROPINFIDT = 0 and ROPINFIDT2= 1"
                Set rsSab = cnsab.Execute(xSQL)
        
                
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect.Col = 0
        If xYROPDOS0.ROPDOSSTA = " " Then
            X = vbCrLf & "          " & dateImp10_S(xYROPDOS0.ROPDOSGECH)
        Else
            X = ""
        End If
        fgSelect.Text = Format$(xYROPDOS0.ROPDOSID, "#0000") & " - " & fgSelect_Display_USR(xYROPDOS0.ROPDOSIUSR) _
                        & vbCrLf & "       => " & fgSelect_Display_USR(xYROPDOS0.ROPDOSGUSR) & X
        fgSelect.Col = 1
        fgSelect.Text = xYROPDOS0.ROPDOSXDOM & vbCrLf & xYROPDOS0.ROPDOSXAPP
        fgSelect.Col = 2
        
        'If Not rsSab.EOF Then
        '    Xdisplay = Replace(rsSab("ROPINFGTXT"), vbCrLf, " | ")
        'Else
        '    Xdisplay = ""
        'End If
       'fgSelect.Text = Xdisplay
        fgSelect.Text = Trim(rsSab("ROPINFGTXT"))
            txtSelect_txt = fgSelect.Text
             HeightOfLine = fgSelect.RowHeightMin / 3 - 20 'Me.TextHeight(txtselect.Text)
    
             LinesOfText = SendMessage(txtSelect_txt.hwnd, EM_GETLINECOUNT, 0&, 0&) + 1
             
             If fgSelect.RowHeight(fgSelect.Row) < (LinesOfText * HeightOfLine) Then
                fgSelect.RowHeight(fgSelect.Row) = LinesOfText * HeightOfLine
             End If

        fgSelect_YROPDOS0_Forecolor
        
    End If
Next K
fgSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Nb de dossiers : " & selYROPDOS0_Nb): DoEvents

fraSelect.Visible = True
cmdPrint.Enabled = True


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
        'For I = fgSelect_arrIndex To fgSelect.FixedCols Step -1
        For I = fgSelect.FixedCols To fgSelect_arrIndex
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
    End If
End If
fgSelect.LeftCol = fgSelect.FixedCols
fgSelect.Visible = True
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
        'For I = fgDetail_arrIndex To fgDetail.FixedCols Step -1
        For I = fgDetail.FixedCols To fgDetail_arrIndex
          fgDetail.Col = I: fgDetail.CellBackColor = lColor
        Next I
    End If
End If
fgDetail.LeftCol = fgDetail.FixedCols
fgDetail.Visible = True
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
    
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub




Public Sub cmdPrint0_Dossier_All()
Dim K As Integer, xSQL As String
On Error Resume Next

prtDROPI_Open 1, "Risques Opérationnels"

For K = 1 To selYROPDOS0_Nb
    
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 where ROPDOSID =" & selYROPDOS0(K).ROPDOSID
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then
        V = rsYROPDOS0_GetBuffer(rsSab, oldYROPDOS0)
        arrYROPINF0_SQL oldYROPDOS0.ROPDOSID
    End If
    fraDossier_STAK
    cmdPrint0_Dossier
Next K
prtDROPI_Close 1
fraDossier.Visible = False
End Sub

Public Sub cmdPrint0_Dossier_All_XDOM()
Dim K As Integer, xSQL As String
On Error Resume Next

prtDROPI_Open 1, "Risques Opérationnels"

For K = 0 To lstW.ListCount - 1
    lstW.ListIndex = K
    xYROPDOS0.ROPDOSID = Mid$(lstW.Text, 25, 10)
    
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 where ROPDOSID =" & xYROPDOS0.ROPDOSID
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then
        V = rsYROPDOS0_GetBuffer(rsSab, oldYROPDOS0)
        arrYROPINF0_SQL oldYROPDOS0.ROPDOSID
    End If
    fraDossier_STAK
    cmdPrint0_Dossier
Next K
prtDROPI_Close 1
fraDossier.Visible = False
End Sub


Public Function cmdSendMail_Txt(lTxt As String, lK As String) As String
Dim lenX As Long, K As Integer, K1 As Integer, K2 As Integer
Dim I As Integer, I1 As Integer, iSpace As Integer
Dim htmlTxt As String, blnEnd As Boolean
Dim wNb As Integer, wReturn As String
Dim blnNext As Boolean

If lK = "H" Then
    wNb = 135
    wReturn = "</NOBR><BR><NOBR>"
    htmlTxt = "<NOBR>" '"<pre>" 'vbCrLf
    'cmdSendMail_Txt = Replace(lTxt, vbCrLf, "<BR>")
    'Exit Function
Else
    wNb = 90 '65
    wReturn = vbCrLf
    htmlTxt = ""
End If

'htmlTxt = ""
K = 1
lenX = Len(lTxt)
blnEnd = False
Do
    K1 = InStr(K, lTxt, vbCrLf)
    If K1 > 0 Then
        blnNext = False
        Do
            If K1 - K < wNb Then
                blnNext = True
                htmlTxt = htmlTxt & Mid$(lTxt, K, K1 - K) & wReturn
            Else
                K2 = K + wNb
                For I = K2 To K2 - 15 Step -1
                    If Mid$(lTxt, I, 1) = " " Then K2 = I: Exit For
                Next I
                htmlTxt = htmlTxt & Mid$(lTxt, K, K2 - K) & wReturn
                K = K2 + 1
            End If
        Loop Until blnNext
        K = K1 + 2
    Else
        blnEnd = True
        K1 = lenX
        blnNext = False
        Do
            If K1 - K < wNb Then
                blnNext = True
                htmlTxt = htmlTxt & Mid$(lTxt, K, K1 - K + 1) & wReturn
            Else
                K2 = K + wNb
                For I = K2 To K2 - 15 Step -1
                    If Mid$(lTxt, I, 1) = " " Then K2 = I: Exit For
                Next I
                htmlTxt = htmlTxt & Mid$(lTxt, K, K2 - K) & wReturn
                K = K2 + 1
            End If
        Loop Until blnNext
         'For I = K To lenX Step wNb
         '   If I + wNb < lenX Then
         '       htmlTxt = htmlTxt & Mid$(lTXT, I, wNb) & wReturn
         '   Else
         '       htmlTxt = htmlTxt & Mid$(lTXT, I, lenX - I + 1)
         '   End If
         'Next I
    End If
Loop Until blnEnd

If lK = "H" Then
    cmdSendMail_Txt = htmlTxt & "</NOBR>" '& "</pre>"
Else
    cmdSendMail_Txt = htmlTxt
End If
End Function

Public Function cmdSendMail_Sta(lSta As String, lSta_backColor As String) As String
'
'
Select Case lSta
    Case " ": cmdSendMail_Sta = "<Font color = #0000FF > " 'en cours"
    Case "F": lSta_backColor = "bgcolor = #E0FFE0": cmdSendMail_Sta = wTD_Txt_ForeColor & "ok"
    Case "A": lSta_backColor = "bgcolor = #F0F0F0": cmdSendMail_Sta = wTD_Txt_ForeColor & "Ann"
    Case Else: lSta_backColor = "bgcolor = #FFFFFF": cmdSendMail_Sta = wTD_Txt_ForeColor & oldYROPDOS0.ROPDOSSTA
End Select
End Function

Public Function cmdSendMail_Ech(lEch As String, lSta As String) As String
If lSta <> " " Then
    cmdSendMail_Ech = "<b>" & wTD_Txt_ForeColor & dateImp10(lEch) & "</b>"
Else
    If lEch > DSys Then
        cmdSendMail_Ech = "<Font color = #0000FF><b>" & dateImp10(lEch) & "</b>"
    Else
        If lEch = DSys Then
            cmdSendMail_Ech = "<Font color = #FF00FF><b>" & dateImp10(lEch) & "</b>"
        Else
            cmdSendMail_Ech = "<Font color = #FF0000><b>" & dateImp10(lEch) & "</b>"
        End If
    End If
End If
End Function


Public Sub lstAut_ROPDOSGUSR_Display(lIndex As Long, lSSIDOMUNIT As String)
Dim K As Integer, K2 As Integer, X As String

lstAut_ROPDOSGUSR.Clear
lIndex = arrROPDOSISRV_ListIndex(Val(Mid$(lSSIDOMUNIT, 2, 2)))
For K = 1 To 76
    X = "S" & Format(K, "00")
    If Mid$(newAut.BIATABTXT, K + 28, 1) <> " " Then
        lstAut_ROPDOSGUSR.AddItem Mid$(newAut.BIATABTXT, K + 28, 1) & " - " & arrROPDOSISRV_Code(K)
    End If
Next K

End Sub


Public Sub cmdSelect_SQL_2_Duplication()
Dim K As Integer, wROPDOSGECH As String, wROPDOSIAMJ As String
Dim blnOk As Boolean

V = dateImp(oldYROPDOS0.ROPDOSIAMJ)
If oldYROPDOS0.ROPDOSGNAT = "M" Then
    mDateDiff_Duplication = DateDiff("m", V, vDsys)
    wROPDOSGECH = DateAdd_AMJ("m", mDateDiff_Duplication, oldYROPDOS0.ROPDOSGECH)
Else
    mDateDiff_Duplication = DateDiff("d", V, vDsys)
    wROPDOSGECH = DateAdd_AMJ("d", mDateDiff_Duplication, oldYROPDOS0.ROPDOSGECH)
End If

wROPDOSIAMJ = DSys

blnDossierReprise = False
blnOk = True
If oldYROPDOS0.ROPDOSID < 10 Then
    If currentROPDOSISRV <> "_S40" Then blnOk = False
End If

If oldYROPDOS0.ROPDOSID = 1 Then
    Select Case currentROPDOSISRV
        Case "_S40":
              wROPDOSGECH = "20071231"
              wROPDOSIAMJ = "20071001"
              oldYROPDOS0.ROPDOSGUSR = "_S40"
              oldYROPDOS0.ROPDOSIUSR = "_S40"
              oldYROPDOS0.ROPDOSGPRV = "V"
              oldYROPDOS0.ROPDOSGNAT = "?"
              arrYROPINF0(1).ROPINFGUSR = "_S40"
              arrYROPINF0(arrYROPINF0_Nb).ROPINFGUSR = "_S40"
                X = InputBox("Indiquez le numéro de dossier (1000-1499)")
                DossierModèle_ROPDOSID = Val(X)
                If DossierModèle_ROPDOSID < 1000 Or DossierModèle_ROPDOSID > 1499 Then
                    MsgBox "N° de dossier invalide"
                    Exit Sub
                Else
                    blnDossierReprise = True
                End If
        Case "_S31":
              wROPDOSGECH = "20071231"
              wROPDOSIAMJ = "20061225"
               X = InputBox("Indiquez le numéro de dossier (1501-1999)")
                DossierModèle_ROPDOSID = Val(X)
                If DossierModèle_ROPDOSID < 1500 Or DossierModèle_ROPDOSID > 1999 Then
                    MsgBox "N° de dossier invalide"
                    Exit Sub
                Else
                    blnDossierReprise = True
                    blnOk = True
                End If
        Case Else
            blnOk = False
    End Select
Else
    oldYROPDOS0.ROPDOSID = 999999999
End If
If Not blnOk Then
    MsgBox "Vous n'êtes pas habilité à utiliser ce type de dossier"
    Exit Sub

Else
    oldYROPDOS0.ROPDOSGECH = wROPDOSGECH
    oldYROPDOS0.ROPDOSIAMJ = wROPDOSIAMJ
    oldYROPDOS0.ROPDOSUVER = 0
    If Trim(oldYROPDOS0.ROPDOSIUSR) = "?" Then
        oldYROPDOS0.ROPDOSIUSR = usrName_UCase
    End If
    ' If Trim(arrYROPINF0(1).ROPINFGUSR) = "?" Then
    '    arrYROPINF0(1).ROPINFGUSR = usrName_UCase
    'End If
    
   oldYROPDOS0.ROPDOSISRV = ""
    If Trim(oldYROPDOS0.ROPDOSXDOM) = "Modèle" Then
        oldYROPDOS0.ROPDOSXDOM = "?"
        oldYROPDOS0.ROPDOSXAPP = ""
    End If
    oldYROPDOS0.ROPDOSUAMJ = DSys
    dupYROPDOS0 = oldYROPDOS0
    xYROPDOS0 = dupYROPDOS0
    
    dupYROPINF0_Nb = arrYROPINF0_Nb
    ReDim dupYROPINF0(dupYROPINF0_Nb)
    For K = 1 To dupYROPINF0_Nb
        dupYROPINF0(K) = arrYROPINF0(K)
         If oldYROPDOS0.ROPDOSGNAT = "M" Then
            dupYROPINF0(K).ROPINFGECH = DateAdd_AMJ("m", mDateDiff_Duplication, arrYROPINF0(K).ROPINFGECH)
         Else
            dupYROPINF0(K).ROPINFGECH = DateAdd_AMJ("d", mDateDiff_Duplication, arrYROPINF0(K).ROPINFGECH)
        End If

        dupYROPINF0(K).ROPINFUVER = 0
        dupYROPINF0(K).ROPINFUAMJ = DSys
    Next K
    dupYROPINF0(1).ROPINFGECH = wROPDOSGECH
    oldYROPINF0 = dupYROPINF0(1)
    xYROPINF0 = oldYROPINF0
    cmdSelect_SQL_2
End If
End Sub

Public Sub cmdUpdate_Init_24_Ok()

'mROPINFSTA_Value = "F": mROPINFSTA_Set = "$": mROPINFSTA_Where = " ": mROPINFSTAK_Set = "V"
mROPINFSTAD_Set = "F": mROPINFSTAK_Set = "V": mROPINFSTAD_Where = " "
newYROPDOS0.ROPDOSSTA = "F"
newYROPDOS0.ROPDOSSTAK = "V"
cmdDossier_Ok_Transaction "Update"

End Sub

Public Sub cmdUpdate_Init_23_Ok()
    
mROPINFSTA_Value = "F": mROPINFSTA_Set = "F": mROPINFSTA_Where = " ": mROPINFSTAK_Set = "V"
cmdDétail_Update_Ok
If xYROPDOS0.ROPDOSSTAK = "V" Then newYROPDOS0 = oldYROPDOS0: cmdUpdate_Init_24_Ok ' fermeture dossier

End Sub

Public Sub lstAut_ROPDOSISRV_Display()
Dim K As Integer
'$JPL_20071004 K = lstAut_ROPDOSISRV.ListIndex + 1
'$JPL_20071004 arrROPDOSISRV_K = CInt(Mid$(arrROPDOSISRV_K1(K), 2, 2)) + 28

arrROPDOSISRV_K = CInt(Mid$(lstAut_ROPDOSISRV.Text, 2, 2)) + 28
Select Case Mid$(newAut.BIATABTXT, arrROPDOSISRV_K, 1)
    Case "H": optAut_ROPDOSISRV_H = True
    Case "R": optAut_ROPDOSISRV_R = True
    Case "D": optAut_ROPDOSISRV_D = True
    Case "C": optAut_ROPDOSISRV_C = True
    Case "X": optAut_ROPDOSISRV_X = True
    Case "I": optAut_ROPDOSISRV_I = True
    Case Else: optAut_ROPDOSISRV_Z = True
    
End Select

End Sub

Public Function cmdSendMail_USR(lK1 As String) As String
Dim K As Integer, X As String

If Mid$(lK1, 1, 1) <> "_" Then
    X = lK1
Else
    K = Val(Mid$(lK1, 3, 2))
    X = arrROPDOSISRV_K1(K) & " : " & arrROPDOSISRV_Lib(K)
End If
cmdSendMail_USR = X
End Function


Public Sub cmdSendMail_Recipient(lUsrName As String)
Dim K1 As Integer
Dim wUsrName As String

blnRecipient = True
wUsrName = Trim(lUsrName)
For K1 = 1 To arrRecipient_Nb
    If arrRecipient(K1) = wUsrName Then blnRecipient = False: Exit Sub
Next K1
If blnRecipient Then
    arrRecipient_Nb = arrRecipient_Nb + 1
    arrRecipient(arrRecipient_Nb) = wUsrName
    If wRecipient <> "" Then wRecipient = wRecipient & ";"
    wRecipient = wRecipient & mailAdresse_Production(lUsrName)
End If

End Sub


Public Sub cmdSelect_SQL_0()
Dim xSQL As String
Me.Enabled = False

oldYROPDOS0.ROPDOSID = 10
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 where ROPDOSID = 10"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    blntxtUpdate_ROPINFGECH_Change = False
    V = rsYROPDOS0_GetBuffer(rsSab, oldYROPDOS0)
    arrYROPINF0_SQL oldYROPDOS0.ROPDOSID
    cmdSelect_SQL_K = "2"
    cmdSelect_SQL_2_Duplication
End If
tabDossier.Tab = 0
fraDossier_cmd.Visible = False
fraDossier_B.Visible = True
fraDossier_B.Enabled = True
Call DTPicker_Set(txtUpdate_ROPDOSGECH, DateAdd_AMJ("d", 7, DSys))

Me.Enabled = True
If txtUpdate_ROPDOSGTXT.Enabled Then txtUpdate_ROPDOSGTXT.SetFocus

End Sub


Public Function cmdUpdate_Init_00_Hab(lGUSR As String) As Boolean
cmdUpdate_Init_00_Hab = False
If usrName_UCase = lGUSR Then
    cmdUpdate_Init_00_Hab = True
Else
    If Mid$(lGUSR, 1, 2) = "_S" Then
        X = Mid$(currentROPDOSISRV_Hab, Val(Mid$(lGUSR, 3, 2)), 1)
        If X = "H" Or X = "R" Or X = "D" Then cmdUpdate_Init_00_Hab = True
    End If
End If

End Function
Public Function cmdSelect_SQL_1_Hab(lGUSR As String) As Boolean
cmdSelect_SQL_1_Hab = False
If Mid$(lGUSR, 1, 2) = "_S" Then
    X = Mid$(currentROPDOSISRV_Hab, Val(Mid$(lGUSR, 3, 2)), 1)
    If X = "H" Or X = "R" Or X = "D" Or X = "C" Or X = "I" Then cmdSelect_SQL_1_Hab = True
End If

End Function


Public Sub lstUpdate_Modèle_Load()
Dim K As Integer
arrYROPDOS0_SQL (" where ROPDOSID > 10 and ROPDOSID < 1000")
For K = 1 To arrYROPDOS0_Nb
    lstUpdate_Modèle.AddItem Format$(arrYROPDOS0(K).ROPDOSID, "0000") & " " & arrYROPDOS0(K).ROPDOSXDOM & " " & arrYROPDOS0(K).ROPDOSXAPP
Next K
lstUpdate_Modèle.Height = 220 * lstUpdate_Modèle.ListCount
End Sub

Private Sub txtUpdate_ROPINFGUO_LostFocus()
txt_LostFocus txtUpdate_ROPINFGUO

End Sub

Private Sub txtUpdate_ROPINFGUSR_Click()
If blnControl Then
    Select Case Mid$(txtUpdate_ROPINFGNAT, 1, 1)
        Case "P": chkUpdate_ROPINFMAIL_P.Caption = fraDossier_Display_USR(txtUpdate_ROPINFGUSR)
        Case "A", "F": chkUpdate_ROPINFMAIL_A.Caption = fraDossier_Display_USR(txtUpdate_ROPINFGUSR)
    
    End Select
End If

End Sub


Public Function fraDétail_Display_PJ_FileName(lTxt, blnDisplay As Boolean) As String
Dim wExtension As String, X As String

wExtension = UCase$(Trim(lTxt))
X = paramROPDOS_Path_DROPI & xYROPINF0.ROPINFID _
    & "\" & xYROPINF0.ROPINFID & "_" & xYROPINF0.ROPINFIDP _
    & "_" & xYROPINF0.ROPINFIDT & "_" & xYROPINF0.ROPINFIDT2 & "." & wExtension
If blnDisplay Then
    If Dir(X) <> "" Then
        Select Case wExtension
         Case "DOC": Call frmElpPrt.WinWord(X)
         Case "XLS": Call frmElpPrt.Excel(X)
         Case "PDF": Call frmElpPrt.Acrord32(X)
         Case "TXT": Call frmElpPrt.WordPad(X) 'NotePad(X)
         Case "RTF": Call frmElpPrt.WordPad(X)
         Case Else: Call frmElpPrt.IExplore(X)
        End Select
    End If
End If
fraDétail_Display_PJ_FileName = X
End Function

Private Sub txtUpdate_ROPINFGUSR_GotFocus()
txt_GotFocus txtUpdate_ROPINFGUSR

End Sub

Private Sub txtUpdate_ROPINFGUSR_LostFocus()
txt_LostFocus txtUpdate_ROPINFGUSR

End Sub



Public Sub fraDossier_Display_Reset()
lblUpdate_ROPDOSGECH.BackColor = fraDossier_C.BackColor '&HC0F0FF
txtUpdate_ROPDOSGECH.ToolTipText = "Echéance à laquelle le dossier doit être clôturer"
lblUpdate_ROPDOSIAMJ.BackColor = fraDossier_C.BackColor '&HC0F0FF
txtUpdate_ROPDOSIAMJ.ToolTipText = "date du constat de l'incident/demande/événement"
txtUpdate_ROPDOSXDOM.BackColor = vbWhite
txtUpdate_ROPDOSXDOM.ToolTipText = "Domaine concerné (X par défaut)"
txtUpdate_ROPDOSXAPP.BackColor = vbWhite
txtUpdate_ROPDOSXAPP.ToolTipText = "Application concernée (X par défaut)"
txtUpdate_ROPDOSGUSR.BackColor = vbWhite
txtUpdate_ROPDOSGUSR.ToolTipText = "Superviseur du Dossier par défaut _S31 : organisation"
txtUpdate_ROPDOSIUSR.BackColor = vbWhite
txtUpdate_ROPDOSIUSR.ToolTipText = "collaborateur initiateur du constat"
txtUpdate_ROPDOSGNAT.BackColor = vbWhite
txtUpdate_ROPDOSGNAT.ToolTipText = "nature du constat : Incident,Demande,Evénement"
txtUpdate_ROPDOSGPRV.BackColor = vbWhite
txtUpdate_ROPDOSGPRV.ToolTipText = "confidentialité : U (privé) |V (service) | W (public)"
txtUpdate_ROPDOSGGRA.BackColor = vbWhite
txtUpdate_ROPDOSGGRA.ToolTipText = "Gravité de l'incident"

txtUpdate_ROPDOSGPRI.ToolTipText = "Priorité de traitement du dossier"

txtUpdate_ROPDOSXID.ToolTipText = "Référence de l'éditeur (CRI, Case,...) "
txtUpdate_ROPDOSIREF.ToolTipText = "Référence interne (Compte, code et n°d'opération,...) "
txtUpdate_ROPDOSGCOU.ToolTipText = "Estimation du coût induit par ce dysfonctionnement en €"

End Sub

Public Sub fraDétail_Display_Reset()
txtUpdate_ROPINFGUSR.BackColor = vbWhite
txtUpdate_ROPINFGUSR.ToolTipText = "Responsable du processus ou de l'action"
txtUpdate_ROPINFGTXT.BackColor = vbWhite
txtUpdate_ROPINFGTXT.ToolTipText = "texte (1024 caractères)"
txtUpdate_ROPINFGUO.BackColor = vbWhite
txtUpdate_ROPINFGUO.ToolTipText = "durée exprimée en hh.mm"
lblUpdate_ROPINFGECH.BackColor = fraUpdate_B.BackColor
txtUpdate_ROPINFGECH.ToolTipText = "Echéance >= jour et =< échéance dossier"

End Sub

Public Sub cmdSelect_SQL_7_Destinataire(lIndex As String)
Dim K1 As Integer, K2 As Integer
Dim xUsr As String, kLen As Integer

If xYROPDOS0.ROPDOSGPRV = "U" Then
    Call cmdSelect_SQL_7_Destinataire_Select(xYROPDOS0.ROPDOSIUSR, lIndex)
Else
    If Mid$(xYROPINF0.ROPINFGUSR, 1, 1) <> "_" Then
        Call cmdSelect_SQL_7_Destinataire_Select(xYROPINF0.ROPINFGUSR, lIndex)
    Else
        K1 = Val(Mid$(xYROPINF0.ROPINFGUSR, 3, 2))
        xUsr = arrROPDOSISRV_Mail(K1)
        kLen = Len(xUsr)
        For K2 = 1 To kLen Step 12
            Call cmdSelect_SQL_7_Destinataire_Select(Mid$(xUsr, K2, 12), lIndex)
        Next K2

    End If
    
End If

End Sub

Public Sub cmdSelect_SQL_7_Destinataire_Select(lUsr As String, lIndex As String)
If blnDestinataire_Select Then
    If Trim(lUsr) = mDestinataire_Select Then lstW.AddItem lUsr & lIndex
Else
    lstW.AddItem lUsr & lIndex
End If

End Sub

Public Sub cmdPrint0_Echéancier_Destinataire()
XPrt.DrawWidth = 8
XPrt.CurrentY = XPrt.CurrentY + 100
XPrt.Line (prtMinX, XPrt.CurrentY + prtlineHeight)-(prtMaxX, XPrt.CurrentY + prtlineHeight), prtLineColor
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight + 20

prtDROPI_Close 2

End Sub
Public Sub cmdSendMail_Echéancier_Destinataire()
wSendMail.Message = wDétail
wSendMail.AsHTML = True
'MsgBox "cmdsensmail Exit"
'Exit Sub
srvSendMail.Monitor wSendMail
lstErr.AddItem "@ " & wSendMail.Recipient
End Sub


Public Sub cmdSendMail_Echéancier_Open(lUsrName As String)
wRecipient = ""
Call cmdSendMail_Recipient(lUsrName)
wSendMail.From = currentSSIWINMAIL
wSendMail.FromDisplayName = usrName_UCase
wSendMail.Recipient = wRecipient
wSendMail.CcRecipient = ""
wSendMail.Attachment = ""

bgColor = "" '"cyan"

wSendMail.Subject = "BIA.RO - " & " Echéancier " & lUsrName & " au " & DSys
wDétail = "<TABLE border = 1  width=1000 height=5 bgcolor=#0000FF cellpadding=4 ><TR>" _
         & "<TD  width=200 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Référence</TD>" _
         & "<TD  width=600 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Responsable</B></TD>" _
         & "<TD  width=100 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Echéance</TD>" _
         & "<TD  width=100 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Statut</TD>" _
        & "</TR></TABLE>"

End Sub

Public Sub cmdSendMail_Echéancier_Dossier()
Dim X As String, xUsr As String, wROPDOSXID As String

wTD_BackColor = "bgcolor = #87CEFA"
wTD_Sta_ForeColor = wTD_BackColor
If Trim(xYROPDOS0.ROPDOSXID) = "" Then
    wROPDOSXID = ""
Else
    wROPDOSXID = " . " & htmlFontColor_Red & xYROPDOS0.ROPDOSXID
End If

X = cmdSendMail_Sta(xYROPDOS0.ROPDOSSTA, wTD_Sta_ForeColor)
xUsr = Trim(cmdSendMail_USR(xYROPDOS0.ROPDOSGUSR)) & htmlFontColor_Red & " . . . " & xYROPDOS0.ROPDOSXDOM & " - " & xYROPDOS0.ROPDOSXAPP
wDétail = wDétail & "<TABLE border = 0  width=1000   cellpadding=4 ><Font  color=#FFFFFF><TR>" _
         & "<TD " & wTD_BackColor & " width=200 height=5><span style='font-size:10.0pt;font-family:Arial'>" & htmlFontColor_Blue & xYROPDOS0.ROPDOSID & wROPDOSXID & "</TD>" _
         & "<TD " & wTD_BackColor & " width=600 height=5><span style='font-size:10.0pt;font-family:Arial'>" & htmlFontColor_Blue & "<B>" & xUsr & "</B></TD>" _
         & "<TD " & wTD_BackColor & " width=100 height=5><span style='font-size:10.0pt;font-family:Arial'>" & cmdSendMail_Ech(xYROPDOS0.ROPDOSGECH, xYROPDOS0.ROPDOSSTA) & "</TD>" _
         & "<TD " & wTD_Sta_ForeColor & " width=100 height=5><span style='font-size:10.0pt;font-family:Arial'>" & X & "</TD>" _
        & "</TR></TABLE>"

End Sub

Public Sub cmdSendMail_Echéancier_Détail()
Dim libROPINFGNAT As String, xUsr As String
If xYROPINF0.ROPINFSTA = " " Then
    If xYROPINF0.ROPINFUAMJ = DSys Then
        wTD_ForeColor = htmlFontColor_Blue
    Else
        wTD_ForeColor = htmlFontColor_Green
    End If
Else
    wTD_ForeColor = htmlFontColor_Gray
End If

Select Case xYROPINF0.ROPINFGNAT
    Case "P":
        wTD_Txt_ForeColor = htmlFontColor_Gray
        libROPINFGNAT = xYROPINF0.ROPINFID & " § " & xYROPINF0.ROPINFIDP
        wTD_BackColor = "bgcolor = #90FFFF"
        wTD_Sta_ForeColor = wTD_BackColor
        X = cmdSendMail_Sta(xYROPINF0.ROPINFSTA, wTD_Sta_ForeColor)
        xUsr = Trim(cmdSendMail_USR(xYROPINF0.ROPINFGUSR))
        wDétail = wDétail & "<NOBR><TABLE  width=1000 border=1   cellpadding=4 ></B><TR>" _
                 & "<TD " & wTD_BackColor & " width=200 height=5><span style='font-size:8.0pt;font-family:Arial'>" & wTD_ForeColor & libROPINFGNAT & "</TD>" _
                 & "<TD " & wTD_BackColor & " width=600 height=5><span style='font-size:8.0pt;font-family:Arial'>" & wTD_ForeColor & xUsr & "</TD>" _
                 & "<TD " & wTD_BackColor & " width=100 height=5><span style='font-size:8.0pt;font-family:Arial'>" & cmdSendMail_Ech(xYROPINF0.ROPINFGECH, xYROPINF0.ROPINFSTA) & "</TD>" _
                 & "<TD " & wTD_Sta_ForeColor & " width=100 height=5><span style='font-size:8.0pt;font-family:Arial'>" & X & "</TD>" _
                 & "</TR>" _
                 & "<TR>" _
                 & "<TD colspan=4  height=5><PRE><span style='font-size:10.0pt;font-family:Arial'>" & wTD_Txt_ForeColor & cmdSendMail_Txt(xYROPINF0.ROPINFGTXT, "H") _
                 & "</TD></TR></TABLE>"
     Case "A", "F":
        wTD_Txt_ForeColor = htmlFontColor_Blue
        libROPINFGNAT = xYROPINF0.ROPINFID & " § " & xYROPINF0.ROPINFIDP & " - " & Format$(xYROPINF0.ROPINFIDT, "00")
     
        wTD_BackColor = "bgcolor =  #B0FFFF"
        wTD_Sta_ForeColor = wTD_BackColor
        X = cmdSendMail_Sta(xYROPINF0.ROPINFSTA, wTD_Sta_ForeColor)

        xUsr = cmdSendMail_USR(xYROPINF0.ROPINFGUSR)
        wDétail = wDétail & "<NOBR><TABLE  width=1000 border=1    cellpadding=4 ></B><TR>" _
                 & "<TD " & wTD_BackColor & " width=200  height=5><span style='font-size:8.0pt;font-family:Arial'>" & wTD_ForeColor & libROPINFGNAT & "</TD>" _
                 & "<TD " & wTD_BackColor & " width=600  height=5><span style='font-size:8.0pt;font-family:Arial'>" & wTD_ForeColor & xUsr & "</TD>" _
                 & "<TD " & wTD_BackColor & " width=100  height=5><span style='font-size:8.0pt;font-family:Arial'>" & cmdSendMail_Ech(xYROPINF0.ROPINFGECH, xYROPINF0.ROPINFSTA) & "</TD>" _
                 & "<TD " & wTD_Sta_ForeColor & " width=100  height=5><span style='font-size:8.0pt;font-family:Arial'>" & X & "</TD>" _
                 & "</TR>" _
                 & "<TR>" _
                 & "<TD colspan=4  height=5><PRE><span style='font-size:10.0pt;font-family:Arial'>" & wTD_Txt_ForeColor & cmdSendMail_Txt(xYROPINF0.ROPINFGTXT, "H") _
                 & "</TD></TR></TABLE>"
End Select
End Sub


Public Sub fraDossier_Display_YROPINF0()
Dim K As Long, X As String
Dim blnEch As Boolean, xEch As String
Dim wEch_Color As Long, wForecolor As Long
Dim HeightOfLine As Long, LinesOfText As Long

fgDetail.Visible = False
fraDossier_STAK
fgDetail.Clear
fgDetail.Rows = 0
'fgDetail.ColWidth(0) = 75
fgDetail.ColWidth(0) = 8800
For K = 1 To arrYROPINF0_Nb
'-------------------------------------------------------------------------
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
        
    fgDetail.Col = 0: fgDetail.Text = K

    xEch = Space$(10)
    xEch = dateImp10(arrYROPINF0(K).ROPINFGECH)
    Select Case arrYROPINF0(K).ROPINFSTAK
        Case "V": wForecolor = RGB(0, 128, 0)
        Case "R": wForecolor = RGB(255, 0, 0)
        Case "O": wForecolor = RGB(255, 128, 0)
        Case "B": wForecolor = vbBlue 'RGB(0, 128, 255)
        Case "A": wForecolor = RGB(128, 128, 128): xEch = "annulé    "
        Case "!": wForecolor = vbMagenta
        Case Else: wForecolor = vbBlue ' RGB(32, 32, 32)
    End Select
    If arrYROPINF0(K).ROPINFGNAT = "N" Then xEch = "(note)    "
    'wForecolor = fgDetail.CellBackColor
    'fgDetail.Col = 0

    If arrYROPINF0(K).ROPINFGNAT = "J" Then
        fgDetail.Text = arrYROPINF0(K).ROPINFGTXT
         If arrYROPINF0(K).ROPINFSTAK = "A" Then
            fgDetail.CellBackColor = RGB(240, 240, 240)
        Else
            fgDetail.CellBackColor = RGB(255, 255, 240)
        End If
        fgDetail.CellForeColor = RGB(192, 64, 0)

    Else
        
        
        fgDetail.CellFontBold = True
        fgDetail.Text = xEch & "    " & UCase(fraDossier_Display_USR(arrYROPINF0(K).ROPINFGUSR)) _
                             & Space$(50) & vbTab & vbTab & "(" & dateImp10(arrYROPINF0(K).ROPINFCAMJ) _
                             & "  " & Trim(LCase(fraDossier_Display_USR(arrYROPINF0(K).ROPINFCUSR))) & ")"
        fgDetail.CellForeColor = wForecolor  'RGB(255, 255, 255)
        Select Case arrYROPINF0(K).ROPINFGNAT
            Case "A", "F":  fgDetail.CellBackColor = RGB(255, 250, 220) 'RGB(112, 112, 255)
            Case "P":    fgDetail.CellBackColor = RGB(255, 230, 200) 'RGB(80, 80, 255)
            Case Else:    fgDetail.CellBackColor = RGB(255, 240, 220)
        End Select
        If arrYROPINF0(K).ROPINFSTAK = "A" Then fgDetail.CellBackColor = RGB(240, 240, 240)
        If arrYROPINF0(K).ROPINFSTAK = "V" Then fgDetail.CellBackColor = RGB(240, 255, 240)
    '_____________________________________________________________________________________
        fgDetail.Col = fgDetail.Cols - 1: fgDetail.Text = K

        fgDetail.Rows = fgDetail.Rows + 1
        fgDetail.Row = fgDetail.Rows - 1
        'fgDetail.Col = 0: fgDetail.Text = K
        fgDetail.Col = 0
        If arrYROPINF0(K).ROPINFSTAK = "A" Then
            fgDetail.CellBackColor = RGB(250, 250, 250)
        Else
            fgDetail.CellBackColor = RGB(255, 255, 255)
        End If
        fgDetail.Text = Trim(arrYROPINF0(K).ROPINFGTXT) '& vbCr
            txtDetail = fgDetail.Text
             HeightOfLine = fgDetail.RowHeightMin - 20 'Me.TextHeight(txtDetail.Text)
    
             LinesOfText = SendMessage(txtDetail.hwnd, EM_GETLINECOUNT, 0&, 0&) + 1
             
             If fgDetail.RowHeight(fgDetail.Row) < (LinesOfText * HeightOfLine) Then
                fgDetail.RowHeight(fgDetail.Row) = LinesOfText * HeightOfLine
             End If
    End If
'-------------------------------------------------------------------------
    If K = 1 Then txtUpdate_ROPDOSGTXT = arrYROPINF0(K).ROPINFGTXT: txtUpdate_ROPDOSGTXT.SelStart = 1
    fgDetail.Col = fgDetail.Cols - 1: fgDetail.Text = K

Next K
fgDetail.Visible = True
currentYROPINF0 = arrYROPINF0(arrYROPINF0_Nb)
End Sub

Public Sub fgSelect_YROPDOS0_Read()
Dim V, xSQL As String
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 where ROPDOSID =" & currentROPDOSID
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    V = rsYROPDOS0_GetBuffer(rsSab, xYROPDOS0)
    oldYROPDOS0 = xYROPDOS0
End If


fraDossier_Display

End Sub

Public Sub cmdUpdate_Exe()
Dim K As Integer
cmdDossier_Ok.Visible = False
cmdDossier_Ok_Close.Visible = False
blnYROPINF0_12X = False
If blnControl Then
    fraUpdate_PJ.Visible = False
    'fraDossier_B.Visible = False
    blnSendMail = True 'False
    If cmdUpdate_K = "00" Then
        
    Else
        blnSelect_Update_EnCours = True
        Select Case cmdUpdate_K
            Case "01": cmdUpdate_Init_01
            Case "02":
                  cmdDossier_Ok_Close.Caption = "Ajouter puis Clôturer cette action"
                   Select Case Mid$(cmdUpdate, 3, 1)
                        Case "F":
                            cmdDossier_Ok_Close.Caption = "Ajouter puis Clôturer ce processus"
                            Call cmdUpdate_Init_02("F")
                        Case "I": blnROPINFIDT_Insérer = True: cmdUpdate_Init_02 ("A")
                        Case Else:  Call cmdUpdate_Init_02("A")
                    End Select
                    
            Case "03": cmdUpdate_Init_03
            Case "05": cmdUpdate_Init_05
            Case "11": cmdUpdate_Init_11
            Case "12":
                        If Mid$(cmdUpdate, 3, 1) = "X" Then
                            blnYROPINF0_12X = True
                            cmdUpdate_Init_11
                        Else
                            cmdUpdate_Init_11
                            If oldYROPINF0.ROPINFGNAT = "F" Then
                                 cmdDossier_Ok_Close.Caption = "Enregistrer + Clôturer ce processus"
                            Else
                                 cmdDossier_Ok_Close.Caption = "Enregistrer + Clôturer cette action"
                            End If
                            If Not blnDossierModèle Then cmdDossier_Ok_Close.Visible = True
                       End If
            Case "13": cmdUpdate_Init_11
                       cmdDossier_Ok_Close.Caption = "Enregistrer + Clôturer ce processus"
                       cmdDossier_Ok_Close.Visible = True
            Case "14": cmdUpdate_Init_14
            Case "21": cmdUpdate_Init_21
            Case "22": cmdUpdate_Init_22
            Case "23": cmdUpdate_Init_23
            Case "24": cmdUpdate_Init_24
            Case "31": cmdUpdate_Init_31
            Case "32": cmdUpdate_Init_32
            Case "33": cmdUpdate_Init_33
            Case "34": cmdUpdate_Init_34
            Case "41": cmdUpdate_Init_41
            Case "42": cmdUpdate_Init_42
            Case "43": cmdUpdate_Init_43
            Case "44": cmdUpdate_Init_44
            Case "51": cmdUpdate_Init_51
            Case "52": cmdUpdate_Init_52
            Case "53": cmdUpdate_Init_53
            Case "54": cmdUpdate_Init_54
            Case "64": cmdUpdate_Init_64
            Case "74": cmdUpdate_Init_74
            Case "00", "10", "20", "30", "40", "50", "60":
            Case Else: MsgBox " NON Géré"
        End Select
    End If
End If

End Sub

Public Sub cmdSendMail_ROPDOSMAIL(lK1 As String)

wRecipient = srvSendMail.Exchange_Distribution("DROPI", lK1)


'$JPL 2013-10-01 Public Sub cmdSendMail_ROPDOSMAIL(lK1 As String)
'$JPL 2013-10-01 Dim X As String, K1 As Integer, K2 As Integer
'$JPL 2013-10-01 Dim rsSabX As New ADODB.Recordset

'$JPL 2013-10-01 X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'ROPDOSMAIL' and BIATABK1 = '" & lK1 & "'"
'$JPL 2013-10-01 Set rsSabX = cnsab.Execute(X)

'$JPL 2013-10-01 If Not rsSabX.EOF Then
'$JPL 2013-10-01     X = Trim(rsSabX("BIATABTXT")) & ";"
'$JPL 2013-10-01     K1 = 1
'$JPL 2013-10-01     Do While K1 < Len(X)
 '$JPL 2013-10-01        K2 = InStr(K1, X, ";")
 '$JPL 2013-10-01        Call cmdSendMail_Recipient(Mid$(X, K1, K2 - K1))
 '$JPL 2013-10-01        K1 = K2 + 1
'$JPL 2013-10-01     Loop
'$JPL 2013-10-01 End If '    If Mid$(x, 7, 1) = "I" Then
'        Call cmdSendMail_Recipient(Trim(rsSabX("BIATABK1")))
'    End If

End Sub

