VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmROPDOS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFFF&
   Caption         =   "Risque Opérationnel"
   ClientHeight    =   10596
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   14748
   Icon            =   "ROPDOS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10596
   ScaleWidth      =   14748
   Begin VB.ListBox lstErr 
      BackColor       =   &H00FFFFFA&
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   8280
      TabIndex        =   4
      Top             =   45
      Width           =   5895
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9975
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   14895
      _ExtentX        =   26268
      _ExtentY        =   17590
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777152
      TabCaption(0)   =   "Risques Opérationnels"
      TabPicture(0)   =   "ROPDOS.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ImageList1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Paramétrage"
      TabPicture(1)   =   "ROPDOS.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraParam"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Habilitations"
      TabPicture(2)   =   "ROPDOS.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraAut"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "."
      TabPicture(3)   =   "ROPDOS.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstW"
      Tab(3).ControlCount=   1
      Begin VB.ListBox lstW 
         BackColor       =   &H80000001&
         Height          =   6768
         Left            =   -74880
         Sorted          =   -1  'True
         TabIndex        =   127
         Top             =   1800
         Visible         =   0   'False
         Width           =   9855
      End
      Begin VB.Frame fraAut 
         BackColor       =   &H00C0F0FF&
         Height          =   9255
         Left            =   -74880
         TabIndex        =   37
         Top             =   480
         Width           =   14535
         Begin VB.ListBox lstAut_Usr 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7800
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   41
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
            TabIndex        =   38
            Top             =   360
            Width           =   8895
            Begin VB.CheckBox chkAut_ROPDOSGUSR_Q 
               BackColor       =   &H00E0FFFF&
               Caption         =   "gestion des qualifications"
               Height          =   255
               Left            =   480
               TabIndex        =   128
               Top             =   1680
               Width           =   2535
            End
            Begin VB.ListBox lstAut_ROPDOSGUSR 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1680
               Left            =   240
               TabIndex        =   98
               Top             =   5760
               Width           =   3105
            End
            Begin VB.Frame fraAut_ROPDOSISRV 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Habilitation utilisateur / service"
               Height          =   3015
               Left            =   240
               TabIndex        =   91
               Top             =   2520
               Width           =   3135
               Begin VB.OptionButton optAut_ROPDOSISRV_Z 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "sans lien avec le service"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   99
                  Top             =   2520
                  Width           =   2415
               End
               Begin VB.OptionButton optAut_ROPDOSISRV_I 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Inspection"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   97
                  Top             =   2160
                  Width           =   2415
               End
               Begin VB.OptionButton optAut_ROPDOSISRV_X 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "X sans habilitation"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   96
                  Top             =   1800
                  Width           =   2415
               End
               Begin VB.OptionButton optAut_ROPDOSISRV_C 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Collaboteur"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   95
                  Top             =   1440
                  Width           =   2415
               End
               Begin VB.OptionButton optAut_ROPDOSISRV_D 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Délégation de gestion"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   94
                  Top             =   1080
                  Width           =   2415
               End
               Begin VB.OptionButton optAut_ROPDOSISRV_R 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Responsable"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   93
                  Top             =   720
                  Width           =   2415
               End
               Begin VB.OptionButton optAut_ROPDOSISRV_H 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Hierarchie"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   92
                  Top             =   360
                  Width           =   2415
               End
            End
            Begin VB.ListBox lstAut_ROPDOSISRV 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4740
               Left            =   3720
               TabIndex        =   90
               Top             =   2520
               Width           =   4905
            End
            Begin VB.CheckBox chkAut_ROPDOSGUSR_H 
               BackColor       =   &H00E0FFFF&
               Caption         =   "gestion des habilitations"
               Height          =   255
               Left            =   480
               TabIndex        =   89
               Top             =   1380
               Width           =   2535
            End
            Begin VB.CheckBox chkAut_ROPDOSGUSR_P 
               BackColor       =   &H00E0FFFF&
               Caption         =   "gestion du paramétrage"
               Height          =   255
               Left            =   480
               TabIndex        =   88
               Top             =   1080
               Width           =   2535
            End
            Begin VB.CheckBox chkAut_ROPINFGUSR 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Responsable d'action"
               Height          =   255
               Left            =   480
               TabIndex        =   43
               Top             =   780
               Width           =   2535
            End
            Begin VB.CheckBox chkAut_ROPDOSGUSR 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Gestionnaire de dossier"
               Height          =   255
               Left            =   480
               TabIndex        =   42
               Top             =   480
               Width           =   2535
            End
            Begin VB.CommandButton cmdAut_Update_Ok 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Enregistrer"
               Height          =   765
               Left            =   7080
               MaskColor       =   &H00E0E0E0&
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   960
               Width           =   1095
            End
            Begin VB.CommandButton cmdAut_Update_Quit 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Abandonner"
               Height          =   765
               Left            =   4440
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label lblAut_ROPDOSGUSR_Mail 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblAut_ROPDOSGUSR_Mail"
               Height          =   255
               Left            =   4440
               TabIndex        =   104
               Top             =   480
               Width           =   3855
            End
         End
      End
      Begin VB.Frame fraParam 
         BackColor       =   &H00C0F0FF&
         Height          =   9375
         Left            =   -74880
         TabIndex        =   26
         Top             =   480
         Width           =   14535
         Begin VB.ListBox lstParam_ROPINFGTXT 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5016
            Left            =   9240
            TabIndex        =   86
            Top             =   960
            Width           =   5055
         End
         Begin VB.Frame fraParam_Update 
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
            Height          =   1935
            Left            =   120
            TabIndex        =   31
            Top             =   7320
            Width           =   14175
            Begin VB.CommandButton cmdParam_Update_Quit 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Abandonner"
               Height          =   765
               Left            =   10320
               Style           =   1  'Graphical
               TabIndex        =   36
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
               TabIndex        =   35
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox txtParam_BIATABTXT 
               Height          =   375
               Left            =   7920
               TabIndex        =   34
               Top             =   480
               Width           =   6015
            End
            Begin VB.TextBox txtParam_BIATABK2 
               Height          =   375
               Left            =   5640
               TabIndex        =   33
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtParam_BIATABK1 
               Height          =   375
               Left            =   120
               TabIndex        =   32
               Top             =   480
               Width           =   1935
            End
         End
         Begin VB.ListBox lstParam_ROPDOSXAPP 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5016
            Left            =   4200
            TabIndex        =   29
            Top             =   960
            Width           =   4695
         End
         Begin VB.ListBox lstParam_ROPDOSXDOM 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5016
            Left            =   240
            TabIndex        =   27
            Top             =   960
            Width           =   3735
         End
         Begin VB.Label lblParam_ROPINFGTXT 
            BackColor       =   &H00C0F0FF&
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
            Height          =   255
            Left            =   10440
            TabIndex        =   87
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblParam_ROPDOSXAPP 
            BackColor       =   &H00C0F0FF&
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
            Height          =   255
            Left            =   5520
            TabIndex        =   30
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblParam_ROPDOSXDOM 
            BackColor       =   &H00C0F0FF&
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
            Height          =   255
            Left            =   1080
            TabIndex        =   28
            Top             =   360
            Width           =   1335
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3600
         Top             =   1800
         _ExtentX        =   995
         _ExtentY        =   995
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   30
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":007C
               Key             =   "Vert"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":0616
               Key             =   "Orange"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":0BB0
               Key             =   "Rouge"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":114A
               Key             =   "D"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":1464
               Key             =   "D_Bleu"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":20B6
               Key             =   "Attention"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":2650
               Key             =   "A"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":296A
               Key             =   "A_Bleu"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":35BC
               Key             =   "A_Orange"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":420E
               Key             =   "A_Rouge"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":4E60
               Key             =   "A_Vert"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":5AB2
               Key             =   "P"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":5DCC
               Key             =   "P_Orange"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":6A1E
               Key             =   "P_Bleu"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":7670
               Key             =   "P_Rouge"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":82C2
               Key             =   "P_Vert"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":8F14
               Key             =   "P_Magenta"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":9B66
               Key             =   "Non"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":9CC0
               Key             =   "Select"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":9FDA
               Key             =   "Stop"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":A42C
               Key             =   "GrandStroumpf"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":AD06
               Key             =   "Ici"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":B020
               Key             =   "Note"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":B8FA
               Key             =   "Trombon1"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":BC14
               Key             =   "Bleu"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":BF66
               Key             =   "F"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":C280
               Key             =   "F_Bleu"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":CED2
               Key             =   "F_Orange"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":DB24
               Key             =   "F_Rouge"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ROPDOS.frx":E776
               Key             =   "F_Vert"
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraTab0 
         BackColor       =   &H00E0FFFF&
         Height          =   9525
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   14640
         Begin VB.Frame fraSelect 
            BackColor       =   &H00E0FFFF&
            Height          =   8175
            Left            =   120
            TabIndex        =   9
            Top             =   1320
            Width           =   14415
            Begin VB.ListBox cmdSelect_Update_G 
               BackColor       =   &H00C0FFC0&
               Height          =   1008
               Left            =   10560
               Sorted          =   -1  'True
               TabIndex        =   139
               Top             =   840
               Width           =   3495
            End
            Begin VB.Frame fraSelect_Update 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   8055
               Left            =   4200
               TabIndex        =   10
               Top             =   120
               Visible         =   0   'False
               Width           =   10095
               Begin VB.TextBox txtUpdate_ROPDOSGTXT 
                  BackColor       =   &H00C0F0FF&
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.4
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   720
                  Left            =   240
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   71
                  Top             =   480
                  Width           =   2055
               End
               Begin VB.ListBox cmdSelect_Update 
                  BackColor       =   &H00C0FFC0&
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.4
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   276
                  Left            =   5880
                  Sorted          =   -1  'True
                  TabIndex        =   130
                  Top             =   840
                  Width           =   3975
               End
               Begin VB.TextBox txtUpdate_ROPINFGTXT 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.4
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   2535
                  Left            =   240
                  MaxLength       =   1024
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   84
                  Top             =   5040
                  Width           =   9615
               End
               Begin VB.CheckBox chkSelect_Update_B 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Afficher dossier"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.4
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   8280
                  TabIndex        =   70
                  Top             =   240
                  Width           =   1695
               End
               Begin VB.CommandButton cmdSelect_Update_Ok 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Enregistrer"
                  Height          =   885
                  Left            =   6960
                  MaskColor       =   &H00E0E0E0&
                  Style           =   1  'Graphical
                  TabIndex        =   23
                  Top             =   2520
                  Width           =   1695
               End
               Begin VB.CommandButton cmdSelect_Update_Quit 
                  BackColor       =   &H008080FF&
                  Caption         =   "Abandonner"
                  Height          =   885
                  Left            =   4080
                  Style           =   1  'Graphical
                  TabIndex        =   20
                  Top             =   2520
                  Width           =   1455
               End
               Begin VB.Frame fraDétail_Update_B 
                  BackColor       =   &H00C0F0FF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   7080
                  Left            =   120
                  TabIndex        =   12
                  Top             =   600
                  Width           =   9855
                  Begin VB.ListBox lstUpdate_Modèle 
                     BackColor       =   &H0080FF80&
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   276
                     Left            =   5760
                     TabIndex        =   119
                     Top             =   720
                     Visible         =   0   'False
                     Width           =   3975
                  End
                  Begin VB.TextBox txtUpdate_ROPINFIDTL 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H000000FF&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   7.8
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   285
                     Left            =   3240
                     TabIndex        =   113
                     Text            =   "123"
                     Top             =   3480
                     Visible         =   0   'False
                     Width           =   615
                  End
                  Begin VB.Frame fraUpdate_ROPINFMAIL 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "Destinataires du mail de suivi"
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1815
                     Left            =   0
                     TabIndex        =   105
                     Top             =   1680
                     Width           =   3855
                     Begin VB.CheckBox chkUpdate_ROPINFMAIL_U 
                        BackColor       =   &H00C0F0FF&
                        Caption         =   "moi"
                        BeginProperty Font 
                           Name            =   "@Arial Unicode MS"
                           Size            =   8.4
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   255
                        Left            =   1440
                        TabIndex        =   110
                        Top             =   1520
                        Width           =   2295
                     End
                     Begin VB.CheckBox chkUpdate_ROPINFMAIL_A 
                        BackColor       =   &H00C0F0FF&
                        Caption         =   "Resp action"
                        BeginProperty Font 
                           Name            =   "@Arial Unicode MS"
                           Size            =   8.4
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   255
                        Left            =   1440
                        TabIndex        =   109
                        Top             =   1200
                        Width           =   2295
                     End
                     Begin VB.CheckBox chkUpdate_ROPINFMAIL_I 
                        BackColor       =   &H00C0F0FF&
                        Caption         =   "Initiateur"
                        BeginProperty Font 
                           Name            =   "@Arial Unicode MS"
                           Size            =   8.4
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   255
                        Left            =   1440
                        TabIndex        =   108
                        Top             =   600
                        Width           =   2295
                     End
                     Begin VB.CheckBox chkUpdate_ROPINFMAIL_P 
                        BackColor       =   &H00C0F0FF&
                        Caption         =   "Resp processus"
                        BeginProperty Font 
                           Name            =   "@Arial Unicode MS"
                           Size            =   8.4
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   255
                        Left            =   1440
                        TabIndex        =   107
                        Top             =   900
                        Width           =   2295
                     End
                     Begin VB.CheckBox chkUpdate_ROPINFMAIL_D 
                        BackColor       =   &H00C0F0FF&
                        Caption         =   "Gest du dossier"
                        BeginProperty Font 
                           Name            =   "@Arial Unicode MS"
                           Size            =   8.4
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00800000&
                        Height          =   255
                        Left            =   1440
                        TabIndex        =   106
                        Top             =   300
                        Width           =   2295
                     End
                     Begin VB.Label lblUpdate_ROPINFMAIL_U 
                        BackColor       =   &H00C0F0FF&
                        Caption         =   "au suivant....."
                        Height          =   255
                        Left            =   120
                        TabIndex        =   124
                        Top             =   1500
                        Width           =   1095
                     End
                     Begin VB.Label lblUpdate_ROPINFMAIL_A 
                        BackColor       =   &H00C0F0FF&
                        Caption         =   "Resp Action"
                        Height          =   255
                        Left            =   120
                        TabIndex        =   123
                        Top             =   1200
                        Width           =   975
                     End
                     Begin VB.Label lblUpdate_ROPINFMAIL_P 
                        BackColor       =   &H00C0F0FF&
                        Caption         =   "Gest Processus"
                        Height          =   255
                        Left            =   120
                        TabIndex        =   122
                        Top             =   900
                        Width           =   1215
                     End
                     Begin VB.Label lblUpdate_ROPINFMAIL_I 
                        BackColor       =   &H00C0F0FF&
                        Caption         =   "Initiateur"
                        Height          =   255
                        Left            =   120
                        TabIndex        =   121
                        Top             =   600
                        Width           =   975
                     End
                     Begin VB.Label lblUpdate_ROPINFMAIL_D 
                        BackColor       =   &H00C0F0FF&
                        Caption         =   "Superviseur"
                        Height          =   255
                        Left            =   120
                        TabIndex        =   120
                        Top             =   300
                        Width           =   1095
                     End
                  End
                  Begin VB.CommandButton cmdSelect_Update_Close 
                     BackColor       =   &H0080FF80&
                     Caption         =   "Enregistrer + fermer processus"
                     Height          =   885
                     Left            =   6840
                     MaskColor       =   &H00E0E0E0&
                     Style           =   1  'Graphical
                     TabIndex        =   102
                     Top             =   3000
                     Width           =   1695
                  End
                  Begin VB.CheckBox chkUpdate_ROPINFGPRV 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "modifiable uniquement par l'auteur"
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Left            =   120
                     TabIndex        =   83
                     Top             =   3720
                     Width           =   3132
                  End
                  Begin VB.TextBox txtUpdate_ROPINFGUO 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Left            =   4560
                     TabIndex        =   24
                     Top             =   1200
                     Width           =   855
                  End
                  Begin VB.ComboBox txtUpdate_ROPINFSTA 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Left            =   4080
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   19
                     Top             =   240
                     Width           =   1335
                  End
                  Begin VB.ComboBox txtUpdate_ROPINFGNAT 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Left            =   960
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   14
                     Top             =   240
                     Width           =   2775
                  End
                  Begin VB.ComboBox txtUpdate_ROPINFGUSR 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Left            =   960
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   13
                     Top             =   720
                     Width           =   4455
                  End
                  Begin MSComCtl2.DTPicker txtUpdate_ROPINFGECH 
                     Height          =   300
                     Left            =   960
                     TabIndex        =   15
                     Top             =   1200
                     Width           =   1335
                     _ExtentX        =   2350
                     _ExtentY        =   529
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
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
                     Format          =   4063235
                     CurrentDate     =   38699.44875
                     MaxDate         =   401768
                     MinDate         =   36526.4425347222
                  End
                  Begin MSComCtl2.DTPicker txtUpdate_ROPINFGECH_Old 
                     Height          =   300
                     Left            =   8280
                     TabIndex        =   126
                     Top             =   3720
                     Visible         =   0   'False
                     Width           =   1335
                     _ExtentX        =   2350
                     _ExtentY        =   529
                     _Version        =   393216
                     CalendarBackColor=   16777215
                     CalendarForeColor=   0
                     CalendarTitleBackColor=   8421504
                     CalendarTitleForeColor=   16777215
                     CalendarTrailingForeColor=   12632256
                     CustomFormat    =   "dd  MM yyy"
                     Format          =   4063235
                     CurrentDate     =   38699.44875
                     MaxDate         =   401768
                     MinDate         =   36526.4425347222
                  End
                  Begin VB.Label lblUpdate_ROPINFIDTL 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "attendre la fin de l'action"
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   9.6
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Left            =   120
                     TabIndex        =   112
                     Top             =   3480
                     Visible         =   0   'False
                     Width           =   2412
                  End
                  Begin VB.Label txtUpdate_ROPINFGTXT_0 
                     BackColor       =   &H00D0FFD0&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "txtUpdate_ROPINFGTXT_0"
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   9.6
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   300
                     Left            =   120
                     TabIndex        =   103
                     Top             =   4080
                     Width           =   9012
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label lblUpdate_ROPINFGUO 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "Durée (HH.MM)"
                     Height          =   255
                     Left            =   3120
                     TabIndex        =   25
                     Top             =   1320
                     Width           =   1215
                  End
                  Begin VB.Label lblUpdate_ROPINFGNAT 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "Nature"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   18
                     Top             =   360
                     Width           =   615
                  End
                  Begin VB.Label lblUpdate_ROPINFGECH 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "Echéance"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   17
                     Top             =   1320
                     Width           =   735
                  End
                  Begin VB.Label lblUpdate_ROPINFGUSR 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "Resp."
                     Height          =   255
                     Left            =   120
                     TabIndex        =   16
                     Top             =   840
                     Width           =   615
                  End
               End
               Begin VB.Label libUpdate_ROPINFID 
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "libUpdate_ROPINFID"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.4
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   118
                  Top             =   7750
                  Width           =   9855
               End
               Begin VB.Label libUpdate_ROPDOSGUSR 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "libUpdate_ROPDOSGUSR"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.4
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   300
                  Left            =   5520
                  TabIndex        =   117
                  Top             =   228
                  Width           =   2772
               End
               Begin VB.Label libUpdate_ROPDOSIUSR 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "libUpdate_ROPDOSIUSR"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.4
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF00FF&
                  Height          =   300
                  Left            =   2400
                  TabIndex        =   116
                  Top             =   240
                  Width           =   3012
               End
               Begin VB.Label libUpdate_ROPDOSID 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "ID"
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.4
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   300
                  Left            =   120
                  TabIndex        =   72
                  Top             =   228
                  Width           =   2172
               End
            End
            Begin VB.ListBox lstUpdate_ROPINFMAIL 
               BackColor       =   &H00D0FFD0&
               Height          =   48
               Left            =   1080
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   125
               Top             =   360
               Width           =   4400
            End
            Begin VB.Frame fraSelect_Update_B 
               BackColor       =   &H00C0F0FF&
               Height          =   7935
               Left            =   -1440
               TabIndex        =   44
               Top             =   1080
               Width           =   4400
               Begin VB.Frame fraSelect_Update_B_G 
                  BackColor       =   &H00C0F0FF&
                  Caption         =   "Caractéristiques"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   4935
                  Left            =   120
                  TabIndex        =   54
                  Top             =   2520
                  Width           =   3975
                  Begin VB.ComboBox txtUpdate_ROPDOSQUAL 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Left            =   1320
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   129
                     Top             =   4440
                     Width           =   2535
                  End
                  Begin VB.ComboBox txtUpdate_ROPDOSIUSR 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Left            =   1320
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   115
                     Top             =   1440
                     Width           =   2535
                  End
                  Begin VB.TextBox txtUpdate_ROPDOSGCOU 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Left            =   1320
                     TabIndex        =   101
                     Top             =   3960
                     Width           =   1215
                  End
                  Begin VB.ComboBox txtUpdate_ROPDOSGNAT 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Left            =   1320
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   60
                     Top             =   2400
                     Width           =   2535
                  End
                  Begin VB.ComboBox txtUpdate_ROPDOSGUSR 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Left            =   1320
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   59
                     Top             =   360
                     Width           =   2535
                  End
                  Begin VB.ComboBox txtUpdate_ROPDOSGPRI 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Left            =   1320
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   58
                     Top             =   3000
                     Width           =   2535
                  End
                  Begin VB.ComboBox txtUpdate_ROPDOSGGRA 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Left            =   1320
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   57
                     Top             =   3480
                     Width           =   2535
                  End
                  Begin VB.ComboBox txtUpdate_ROPDOSGPRV 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Left            =   2640
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   56
                     Top             =   1920
                     Width           =   1215
                  End
                  Begin VB.ComboBox txtUpdate_ROPDOSSTA 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Left            =   2880
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   55
                     Top             =   960
                     Width           =   975
                  End
                  Begin MSComCtl2.DTPicker txtUpdate_ROPDOSGECH 
                     Height          =   300
                     Left            =   1320
                     TabIndex        =   61
                     Top             =   960
                     Width           =   1215
                     _ExtentX        =   2138
                     _ExtentY        =   529
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
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
                     Format          =   4063235
                     CurrentDate     =   38699.44875
                     MaxDate         =   401768
                     MinDate         =   36526.4425347222
                  End
                  Begin MSComCtl2.DTPicker txtUpdate_ROPDOSIAMJ 
                     Height          =   345
                     Left            =   1320
                     TabIndex        =   62
                     Top             =   1920
                     Width           =   1335
                     _ExtentX        =   2350
                     _ExtentY        =   614
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
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
                     Format          =   4063235
                     CurrentDate     =   38699.44875
                     MaxDate         =   401768
                     MinDate         =   36526.4425347222
                  End
                  Begin VB.Label lblUpdate_ROPDOSIUSR 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "Initiateur"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   114
                     Top             =   1560
                     Width           =   975
                  End
                  Begin VB.Label lblUpdate_ROPDOSGCOU 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "Coût €"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   100
                     Top             =   4080
                     Width           =   1095
                  End
                  Begin VB.Label lblUpdate_ROPDOSSTA 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "état"
                     Height          =   255
                     Left            =   2520
                     TabIndex        =   73
                     Top             =   1080
                     Width           =   375
                  End
                  Begin VB.Label lblUpdate_ROPDOSGNAT 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "Nature"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   69
                     Top             =   2520
                     Width           =   615
                  End
                  Begin VB.Label lblUpdate_ROPDOSGUSR 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "Superviseur"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   68
                     Top             =   480
                     Width           =   975
                  End
                  Begin VB.Label lblUpdate_ROPDOSGPRI 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "Priorité"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   67
                     Top             =   3000
                     Width           =   615
                  End
                  Begin VB.Label lblUpdate_ROPDOSGECH 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "Echéance"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   66
                     Top             =   1080
                     Width           =   855
                  End
                  Begin VB.Label lblUpdate_ROPDOSGGRA 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "Gravité"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   65
                     Top             =   3480
                     Width           =   735
                  End
                  Begin VB.Label lblUpdate_ROPDOSQUAL 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "Qualification"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   64
                     Top             =   4440
                     Width           =   1095
                  End
                  Begin VB.Label lblUpdate_ROPDOSIAMJ 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "date du constat"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   63
                     Top             =   2040
                     Width           =   1215
                  End
               End
               Begin VB.Frame fraSelect_Update_B_X 
                  BackColor       =   &H00C0F0FF&
                  Caption         =   "Références du dossier"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2295
                  Left            =   120
                  TabIndex        =   45
                  Top             =   240
                  Width           =   3975
                  Begin VB.TextBox txtUpdate_ROPDOSXID 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Left            =   1320
                     TabIndex        =   49
                     Top             =   1800
                     Width           =   2535
                  End
                  Begin VB.TextBox txtUpdate_ROPDOSIREF 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Left            =   1320
                     TabIndex        =   48
                     Top             =   1320
                     Width           =   2535
                  End
                  Begin VB.ComboBox txtUpdate_ROPDOSXAPP 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Left            =   120
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   47
                     Top             =   960
                     Width           =   3735
                  End
                  Begin VB.ComboBox txtUpdate_ROPDOSXDOM 
                     BeginProperty Font 
                        Name            =   "@Arial Unicode MS"
                        Size            =   8.4
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Left            =   120
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   46
                     Top             =   480
                     Width           =   3735
                  End
                  Begin VB.Label lblUpdate_ROPDOSXID 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "Réf externe"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   53
                     Top             =   1900
                     Width           =   855
                  End
                  Begin VB.Label lblUpdate_ROPDOSIREF 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "Réf interne"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   52
                     Top             =   1440
                     Width           =   975
                  End
                  Begin VB.Label lblUpdate_ROPDOSXAPP 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "/  Application"
                     Height          =   255
                     Left            =   960
                     TabIndex        =   51
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.Label lblUpdate_ROPDOSXDOM 
                     BackColor       =   &H00C0F0FF&
                     Caption         =   "Domaine"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   50
                     Top             =   240
                     Width           =   735
                  End
               End
               Begin VB.Label libUpdate_ROPDOSUUSR 
                  BackColor       =   &H00C0F0FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "saisie"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   74
                  Top             =   7560
                  Width           =   3735
               End
            End
            Begin VB.ListBox libUpdate_ROPINFGTXT 
               BackColor       =   &H00D0FFD0&
               Height          =   240
               Left            =   1680
               TabIndex        =   85
               Top             =   1200
               Width           =   4400
            End
            Begin VB.Frame fraDétail_Update_J 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Pièce Jointe"
               Height          =   7935
               Left            =   600
               TabIndex        =   75
               Top             =   2280
               Visible         =   0   'False
               Width           =   4400
               Begin VB.FileListBox filDoc 
                  ForeColor       =   &H00008000&
                  Height          =   3528
                  Left            =   120
                  Pattern         =   "*.doc;*.pdf;*.rtf;*.xls;*.txt"
                  TabIndex        =   78
                  Top             =   3240
                  Width           =   4000
               End
               Begin VB.DirListBox dirListBox 
                  Height          =   2340
                  Left            =   240
                  TabIndex        =   77
                  Top             =   720
                  Width           =   4000
               End
               Begin VB.DriveListBox DriveListBox 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   76
                  Top             =   360
                  Width           =   4000
               End
            End
            Begin MSComctlLib.TreeView tvwSelect 
               Height          =   7935
               Left            =   0
               TabIndex        =   11
               Top             =   120
               Width           =   4215
               _ExtentX        =   7430
               _ExtentY        =   13991
               _Version        =   393217
               Indentation     =   882
               LabelEdit       =   1
               LineStyle       =   1
               Sorted          =   -1  'True
               Style           =   7
               ImageList       =   "ImageList1"
               BorderStyle     =   1
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.CheckBox chkUpdate_ROPINFMAIL 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Suivi mail automatique"
            Height          =   255
            Left            =   10440
            TabIndex        =   111
            Top             =   960
            Width           =   1935
         End
         Begin VB.ComboBox cboSelect_SQL 
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   10440
            Sorted          =   -1  'True
            TabIndex        =   8
            Text            =   "cboSelect_SQL"
            Top             =   240
            Width           =   3975
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Rechercher"
            Height          =   645
            Left            =   12600
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   650
            Width           =   1815
         End
         Begin VB.Frame fraSelect_Options_1 
            BackColor       =   &H00E0FFFF&
            Height          =   1440
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   10155
            Begin VB.ComboBox txtSelect_ROPDOSGPRV 
               Height          =   315
               Left            =   2640
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   138
               Top             =   840
               Width           =   1335
            End
            Begin VB.CheckBox chkSelect_ROPDOSGUSR 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Resp"
               Height          =   255
               Left            =   120
               TabIndex        =   137
               Top             =   840
               Width           =   735
            End
            Begin VB.TextBox txtSelect_ROPDOSGUSR 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   960
               TabIndex        =   136
               Top             =   840
               Width           =   1335
            End
            Begin VB.CheckBox chkSelect_ROPDOSGECH 
               BackColor       =   &H00E0FFFF&
               Caption         =   "éch <="
               Height          =   255
               Left            =   120
               TabIndex        =   134
               Top             =   360
               Width           =   855
            End
            Begin VB.ComboBox txtSelect_ROPDOSSTA 
               Height          =   315
               Left            =   2640
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   133
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtSelect_ROPINFGTXT 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8880
               TabIndex        =   131
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox txtSelect_ROPDOSXID 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8880
               TabIndex        =   82
               Top             =   600
               Width           =   1095
            End
            Begin VB.ComboBox txtSelect_ROPDOSXAPP 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   4080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   80
               Top             =   840
               Width           =   3135
            End
            Begin VB.ComboBox txtSelect_ROPDOSXDOM 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   4080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   79
               Top             =   360
               Width           =   3135
            End
            Begin VB.TextBox txtSelect_ROPDOSID 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8880
               TabIndex        =   22
               Top             =   240
               Width           =   1095
            End
            Begin MSComCtl2.DTPicker txtSelect_ROPDOSGECH_Max 
               Height          =   300
               Left            =   960
               TabIndex        =   135
               Top             =   360
               Width           =   1335
               _ExtentX        =   2350
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.4
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
               Format          =   4063235
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_ROPINFGTXT 
               BackColor       =   &H00E0FFFF&
               Caption         =   "texte à rechercher"
               Height          =   375
               Left            =   7680
               TabIndex        =   132
               Top             =   960
               Width           =   855
            End
            Begin VB.Label lblSelect_ROPDOSXID 
               BackColor       =   &H00E0FFFF&
               Caption         =   "N° CRI"
               Height          =   255
               Left            =   7680
               TabIndex        =   81
               Top             =   600
               Width           =   735
            End
            Begin VB.Label lblSelect_ROPDOSID 
               BackColor       =   &H00E0FFFF&
               Caption         =   "N° dossier"
               Height          =   255
               Left            =   7680
               TabIndex        =   21
               Top             =   240
               Width           =   975
            End
         End
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   14280
      Picture         =   "ROPDOS.frx":F3C8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   500
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
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
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
   Begin VB.Menu mnuParam 
      Caption         =   "mnuParam"
      Visible         =   0   'False
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
      Begin VB.Menu mnuPrint0_Dossier_All 
         Caption         =   "Imprimer tous les dossiers"
      End
   End
End
Attribute VB_Name = "frmROPDOS"
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
Dim ROPDOS_Aut As typeAuthorization, mAPP_Menu As String
Dim blnTransaction As Boolean
Dim blnAuto As Boolean, blnAuto_Ok As Boolean
Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
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
'______________________________________________________________________

Dim fgDétail_FormatString As String, fgDétail_K As Integer
Dim fgDétail_RowDisplay As Integer, fgDétail_RowClick As Integer, fgDétail_ColClick As Integer
Dim fgDétail_ColorClick As Long, fgDétail_ColorDisplay As Long
Dim fgDétail_Sort1 As Integer, fgDétail_Sort2 As Integer
Dim fgDétail_SortAD As Integer, fgDétail_Sort1_Old As Integer
Dim fgDétail_arrIndex As Integer
Dim blnfgDétail_DisplayLine As Boolean
Dim cmdDétail_Ok_Caption As String
Dim cmdDétail_SQL_K As String
Dim xYROPINF0 As typeYROPINF0, meYROPINF0 As typeYROPINF0
Dim newYROPINF0 As typeYROPINF0, oldYROPINF0 As typeYROPINF0
Dim arrYROPINF0() As typeYROPINF0, arrYROPINF0_Nb As Long, arrYROPINF0_Max As Long, arrYROPINF0_Index As Long
Dim selYROPINF0() As typeYROPINF0
Dim zYROPINF0 As typeYROPINF0
Dim mailYROPINF0 As typeYROPINF0, mailYROPINF0_Suivant As typeYROPINF0
Dim mUpdate_Nature As String, mUpdate_Action As String
'______________________________________________________________________

Dim xNode As MSComctlLib.Node, xNode_Parent As MSComctlLib.Node, mSelect_Node As MSComctlLib.Node
Dim blnmSelect_Node As Boolean
'Dim mSelect_Node_Key As String, mUpdate_Node_Key As String
Dim cmdSelect_Update_K As String, cmdSelect_Update_Init_K As String, cmdSelect_Update_Fct As String
Dim blntvwSelect_NodeClick As Boolean, mROPINFSTA_Value As String, mROPINFSTA_Set As String, mROPINFSTA_Where As String
Dim mROPINFSTAK_Set As String

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
Dim blnSelect_Update_EnCours As Boolean, blnSelect_Update_Expandable As Boolean
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
Dim blntvwSelect_Click As Boolean
Dim fraSelect_Update_Left As Long, fraSelect_Update_Right As Long

Dim blnSelect_Update_B_Display As Boolean
Dim blntxtUpdate_ROPINFGECH_Change As Boolean

Dim blnROPDOSQUAL As Boolean
'___________________________________________________________________________
Dim wTD_BackColor As String, wTD_ForeColor As String, wTD_Sta_ForeColor As String, wTD_Txt_ForeColor As String
Dim htmlFontColor_Blue As String, htmlFontColor_Green As String, htmlFontColor_Gray As String, htmlFontColor_Red As String
Dim paramROPDOS_Path_DROPI As String
Private Sub arrYROPDOS0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrYROPDOS0(501)
arrYROPDOS0_Max = 500: arrYROPDOS0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYROPDOS0_GetBuffer(rsSab, xYROPDOS0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmROPDOS.fgselect_Display"
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
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrYROPINF0(501)
arrYROPINF0_Max = 500: arrYROPINF0_Nb = 0
rsYROPINF0_Init arrYROPINF0(0) ': Processus_Index = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YROPINF0" _
     & " where ROPINFID = " & lROPDOSID _
     & " order by ROPINFIDP,ROPINFIDT,ROPINFIDT2"
     
Set rsSab = cnsab.Execute(xSql)

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
Dim I As Long, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
tvwSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_1"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = True
fraSelect_Options_1.Enabled = True
chkSelect_ROPDOSGECH.Enabled = True
txtSelect_ROPDOSSTA.Enabled = True
txtSelect_ROPDOSGUSR.Enabled = False
chkSelect_ROPDOSGUSR = "1"
chkSelect_ROPDOSGUSR.Enabled = False
fraSelect_Options_1.Visible = True
Select Case cmdSelect_SQL_K
    Case "1": txtSelect_ROPDOSGUSR.Text = "" 'usrName_UCase & " + " & currentROPDOSISRV_Nom
    Case "1U": txtSelect_ROPDOSGUSR.Text = usrName_UCase:
    Case "1V": txtSelect_ROPDOSGUSR.Text = currentROPDOSISRV_Nom
    Case "1W": txtSelect_ROPDOSGUSR.Text = ""

End Select

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub

Private Sub lstSelect_Load_2()
Dim I As Long, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
tvwSelect.Visible = False
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
Dim I As Long, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
tvwSelect.Visible = False
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
Dim I As Long, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
tvwSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_7"
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

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
Select Case SSTab1.Tab
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
        If fraDétail_Update_J.Visible Then
            fraDétail_Update_J.Visible = False
            Exit Sub
        End If
        If libUpdate_ROPINFGTXT.Visible Then
            libUpdate_ROPINFGTXT.Visible = False
            Exit Sub
        End If
        If lstUpdate_ROPINFMAIL.Visible Then
            lstUpdate_ROPINFMAIL.Visible = False
            Exit Sub
        End If
        If fraSelect_Update_B.Visible And cmdSelect_Update.Height > 300 Then
            cmdSelect_Update.Height = 220
            Exit Sub
        End If
        If blnSelect_Update_EnCours Then
            cmdSelect_Update_Reset
            Exit Sub
        End If
        fraSelect_Update_B.Visible = False
        If fraSelect_Update.Visible Then
            fraSelect_Update.Visible = False: cmdSelect_Update_Ok.Visible = False: cmdSelect_Update_Close.Visible = False
            fraDétail_Update_J.Visible = False
            cmdSelect_Update_G.Visible = False
            cmdPrint.Enabled = True
            Exit Sub
        End If
        If tvwSelect.Visible Then tvwSelect.Visible = False: cmdSelect_Ok.Caption = "Extraire les factures": Exit Sub
        Unload Me
End Select
End Sub





Private Sub cboSelect_SQL_Click()
cmdSelect_SQL_K = Trim(Mid$(cboSelect_SQL, 1, 2))
If blnControl Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    tvwSelect.Visible = False
    fraSelect_Update_B.Visible = False
    libUpdate_ROPINFGTXT.Visible = False
    lstUpdate_ROPINFMAIL.Visible = False
    fraSelect_Options_1.Visible = False
    fraSelect_Update.Visible = False
    Select Case cmdSelect_SQL_K
        Case "0": cmdSelect_Ok_Click
        Case "1", "1U", "1V", "1W", "1X": lstSelect_Load_1
        Case "1M": cmdSelect_Ok_Click
        Case "2": lstSelect_Load_2
        Case "2M": lstSelect_Load_2
        Case "6": lstSelect_Load_6
        Case "7": lstSelect_Load_7
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

Private Sub chkSelect_ROPDOSGUSR_Click()
If chkSelect_ROPDOSGUSR = "1" Then
    txtSelect_ROPDOSGUSR.Visible = True
Else
    txtSelect_ROPDOSGUSR.Visible = False
End If

End Sub


Private Sub chkSelect_Update_B_Click()
If chkSelect_Update_B.Value = "1" Then
    If blnSelect_Update_B_Display Then
        fraSelect_Update_B.Visible = True
        txtUpdate_ROPDOSGTXT.Visible = False: cmdSelect_Update_G.Visible = False
    Else
        If cmdSelect_SQL_X1 = "1" Then txtUpdate_ROPDOSGTXT.Visible = True: cmdSelect_Update_G.Visible = True: fraSelect_Update.Left = fraSelect_Update_Left
    End If
Else
    fraSelect_Update_B.Visible = False
    txtUpdate_ROPDOSGTXT.Visible = False: cmdSelect_Update_G.Visible = False
End If
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
Dim K As Integer

Me.Enabled = False: Me.MousePointer = vbHourglass
'___________________________________________________________
App_Debug = "cmdAut_Update_Ok"
blnOk = False

K = InStr(28, newAut.BIATABTXT, "R")
If K > 0 Then
    blnOk = True
Else
    K = InStr(28, newAut.BIATABTXT, "D")
    If K > 0 Then
        blnOk = True
    Else
        K = InStr(28, newAut.BIATABTXT, "C")
        If K > 0 Then
            blnOk = True
        Else
            K = InStr(28, newAut.BIATABTXT, "X")
            If K > 0 Then blnOk = True
        End If
    End If
End If

If Not blnOk Then
    Call lstErr_Clear(lstErr, cmdContext, "? préciser le service de ce collaborateur"): DoEvents
Else
    Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement " & App_Debug): DoEvents
    Mid$(newAut.BIATABTXT, 26, 3) = "S" & Format$(K - 28, "00")
    cmdAut_Update_Ok_Transaction
    
    Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement" & App_Debug): DoEvents
End If
'___________________________________________________________
Exit_Sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdAut_Update_Quit_Click()
fraAut_Update.Visible = False

End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdSelect_Update_Ok_04()
Dim V

App_Debug = "cmdSelect_Update_Ok_04"
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement " & App_Debug): DoEvents
'$20071122_JPL$ chkSelect_Update_B.Value = "1"

If IsNull(fraSelect_Update_Control) Then
    oldYROPDOS0 = newYROPDOS0
    If IsNull(fraDétail_Update_Control) Then
    
    'Création Dossier : insert YROPDOS0 et YROPINF0 (description)
    '---------------------------------------------------------------

        V = cmdSelect_Update_Ok_Transaction("Insert")
        If Not IsNull(V) Then
            MsgBox V, vbCritical, Me.Name & " : cmddétail_Update_Ok" & App_Debug
            Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
        Else
            If blnSendMail Then Call cmdSendMail("D")
            cmdSelect_SQL_K = 1: cmdSelect_SQL_X1 = 1
            Call cmdSelect_SQL_1(newYROPDOS0.ROPDOSID)
            Call cmdSelect_Update_Reset
        End If
    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement" & App_Debug): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Function cmdSelect_Update_Ok_Transaction(lFct As String)
Dim V, X As String, xSql As String
Dim xSet As String, xWhere As String
Dim Nb As Long, K As Integer
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdSelect_Update_Ok_Transaction"
'-------------------------------------------------------
cmdSelect_Update_Ok_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

'________________________________________________________________________________
Call cmdSelect_Update_Ok_GSRV(newYROPDOS0.ROPDOSIUSR, newYROPDOS0.ROPDOSISRV)
Call cmdSelect_Update_Ok_GSRV(newYROPDOS0.ROPDOSGUSR, newYROPDOS0.ROPDOSGSRV)
mailYROPDOS0 = newYROPDOS0: mailYROPINF0 = newYROPINF0

Select Case lFct
    Case "Update": V = sqlYROPDOS0_Update(newYROPDOS0, oldYROPDOS0)
    Case "Insert":
                If blnDossierModèle Or blnDossierReprise Then
                    newYROPDOS0.ROPDOSID = DossierModèle_ROPDOSID
                Else
                    V = sqlROPDOSID_Init("ROPDOSID", newYROPDOS0.ROPDOSID)
                End If
                If IsNull(V) Then
                    Call cmdSelect_Update_Ok_GSRV(newYROPINF0.ROPINFGUSR, newYROPINF0.ROPINFGSRV)
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
                            Call cmdSelect_Update_Ok_GSRV(newYROPINF0.ROPINFGUSR, newYROPINF0.ROPINFGSRV)
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
Select Case cmdSelect_Update_K
    Case 14
        newYROPINF0 = arrYROPINF0(1)
        newYROPINF0.ROPINFGECH = newYROPDOS0.ROPDOSGECH
        V = sqlYROPINF0_Update(newYROPINF0, arrYROPINF0(1))
        If Not IsNull(V) Then GoTo Error_MsgBox
        arrYROPINF0(1) = newYROPINF0

    Case 24, 34, 54:
        xSet = " set ROPINFSTA = '" & mROPINFSTA_Set & "'" & " , ROPINFSTAK = '" & mROPINFSTAK_Set & "'"
        
        xWhere = " where ROPINFID = " & oldYROPDOS0.ROPDOSID _
       & " and ROPINFSTA = '" & mROPINFSTA_Where & "'"

        V = sqlYROPINF0_Requête("update ", xSet, xWhere)
        If Not IsNull(V) Then GoTo Error_MsgBox
    Case 44:
        
        xWhere = " where ROPINFID = " & oldYROPDOS0.ROPDOSID
       
        V = sqlYROPINF0_Requête("delete from ", "", xWhere)
        If Not IsNull(V) Then GoTo Error_MsgBox
End Select
'________________________________________________________________________________

GoTo Exit_Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_Sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        cmdSelect_Update_Reset
    End If
    
    cmdSelect_Update_Ok_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function
Public Function cmdSelect_Update_Ok_Transaction_Duplication()
Dim V, X As String, xSql As String
Dim xSet As String, xWhere As String
Dim Nb As Long, K As Integer
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdSelect_Update_Ok_Transaction_Duplication"
'-------------------------------------------------------
cmdSelect_Update_Ok_Transaction_Duplication = Null
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

GoTo Exit_Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_Sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        If blnSendMail Then Call cmdSendMail("D")
        cmdSelect_SQL_K = 1: cmdSelect_SQL_X1 = 1
        Call cmdSelect_SQL_1(newYROPDOS0.ROPDOSID)

    End If
    
    cmdSelect_Update_Ok_Transaction_Duplication = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function

Public Function cmdParam_Update_Ok_Transaction()
Dim V, X As String, xSql As String
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
    Case "Insert":
                If Trim(oldParam.BIATABID) = "ROPINFGTXT" Then
                    V = sqlROPDOSID_Init("ROPINFGTXT_$", wseq)
                    newParam.BIATABK2 = Format$(wseq, "000000000000")
                End If
                If IsNull(V) Then
                    V = sqlYBIATAB0_Insert(newParam)    ' dossier & description
                End If
    Case "Delete": V = sqlYBIATAB0_Delete(oldParam)
    Case Else: V = "? fct non traitée : " & cmdParam_SQL_K
End Select
If Not IsNull(V) Then GoTo Error_MsgBox

Select Case Trim(oldParam.BIATABID)
    Case "ROPDOSXDOM": lstParam_ROPDOSXDOM_Load: sqlYBIATAB0_cboID "ROPDOSXDOM", txtUpdate_ROPDOSXDOM

    Case "ROPDOSXAPP": lstParam_ROPDOSXAPP_Load (oldParam.BIATABK1)
    Case "ROPINFGTXT": lstParam_ROPINFGTXT_Load
End Select

'________________________________________________________________________________

GoTo Exit_Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_Sub:
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
Dim V, X As String, xSql As String
Dim xSet As String, xWhere As String
Dim Nb As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdAut_Update_Ok_Transaction"
'-------------------------------------------------------
cmdAut_Update_Ok_Transaction = Null

Mid$(newAut.BIATABTXT, 1, 4) = "****"
If chkAut_ROPDOSGUSR = "1" Then Mid$(newAut.BIATABTXT, 1, 1) = "D"
If chkAut_ROPINFGUSR = "1" Then Mid$(newAut.BIATABTXT, 2, 1) = "A"
If chkAut_ROPDOSGUSR_P = "1" Then Mid$(newAut.BIATABTXT, 3, 1) = "P"
If chkAut_ROPDOSGUSR_H = "1" Then Mid$(newAut.BIATABTXT, 4, 1) = "H"
If chkAut_ROPDOSGUSR_Q = "1" Then Mid$(newAut.BIATABTXT, 5, 1) = "Q"

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

GoTo Exit_Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_Sub:
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
Dim V, X As String, xSql As String
Dim xSet As String, xWhere As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdDétail_Update_Ok_Transaction"
'-------------------------------------------------------
cmdDétail_Update_Ok_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
mailYROPDOS0 = oldYROPDOS0: mailYROPINF0 = newYROPINF0
If cmdSelect_Update_K = "05" Then
    newYROPINF0.ROPINFGUSR = newFileExtension
Else
    Call cmdSelect_Update_Ok_GSRV(newYROPINF0.ROPINFGUSR, newYROPINF0.ROPINFGSRV)
End If
Select Case lFct
    Case "Update":
            V = sqlYROPINF0_Update(newYROPINF0, oldYROPINF0)
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
Select Case cmdSelect_Update_K
    Case "05"
        newFileName = newDirPath & "\" & newYROPINF0.ROPINFID & "_" & newYROPINF0.ROPINFIDP _
            & "_" & newYROPINF0.ROPINFIDT & "_" & newYROPINF0.ROPINFIDT2 & "." & newFileExtension

        If Not msFileSystem.FolderExists(newDirPath) Then MkDir newDirPath
        msFileSystem.CopyFile oldFileName, newFileName

    Case "11", "12", "13"
        newYROPDOS0 = oldYROPDOS0
        '$JPL_20071004 newYROPDOS0.ROPDOSIUSR = newYROPINF0.ROPINFGUSR
        cmdSelect_Update_Ok_04_ISRV
        V = sqlYROPDOS0_Update(newYROPDOS0, oldYROPDOS0)
        mailYROPDOS0 = newYROPDOS0
        If Not IsNull(V) Then GoTo Error_MsgBox
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


GoTo Exit_Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_Sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        cmdSelect_Update_Reset

    End If
    
    cmdDétail_Update_Ok_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

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
'________________________________________________________________________________
    For K = 1 To arrYROPINF0_Nb
        If selYROPINF0(K).ROPINFSTAK <> arrYROPINF0(K).ROPINFSTAK Then
            V = sqlYROPINF0_Update(selYROPINF0(K), arrYROPINF0(K))
        End If
    Next K
'________________________________________________________________________________

GoTo Exit_Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_Sub:
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
App_Debug = "cmdSelect_Update_Ok"

If IsNull(fraDétail_Update_Control) Then

    V = cmdDétail_Update_Ok_Transaction(cmdSelect_Update_Fct)
    If Not IsNull(V) Then
        MsgBox V, vbCritical, Me.Name & " : cmddétail_Update_Ok" & App_Debug
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    Else
        If blntvwSelect_NodeClick Then
            If Mid$(cmdSelect_Update_K, 1, 1) <> "4" Then blnmSelect_Node = True 'se repositionner sauf si effacement
            
            tvwSelect_NodeClick tvwSelect.Nodes("D" & Format$(oldYROPDOS0.ROPDOSID, "000000000"))
        Else
            arrYROPINF0(arrYROPINF0_Index) = newYROPINF0
            tvwSelect_Display
        End If
        If blnSendMail Then Call cmdSendMail(" ")
        cmdSelect_Update_Reset

    End If
End If
End Sub
Private Sub cmdDétail_Update_Ok_Insert()
Dim V, X As String, K As Integer
App_Debug = "cmdDétail_Update_Ok_Insert"

If IsNull(fraDétail_Update_Control) Then

    Select Case cmdSelect_Update_K
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
        blnmSelect_Node = True
        tvwSelect_NodeClick tvwSelect.Nodes("D" & Format$(oldYROPDOS0.ROPDOSID, "000000000"))
        If blnSendMail Then Call cmdSendMail(" ")
        cmdSelect_Update_Reset
    End If
End If

End Sub

Private Sub cmdSelect_Update_Ok_14()
Dim V, X As String
App_Debug = "cmdSelect_Update_Ok"

    If IsNull(fraSelect_Update_Control) Then
    
        V = cmdSelect_Update_Ok_Transaction("Update")
        If Not IsNull(V) Then
            MsgBox V, vbCritical, Me.Name & " : cmdSelect_Update_Ok" & App_Debug
            Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
        Else
            oldYROPDOS0 = newYROPDOS0
            tvwSelect_Display
            If blnSendMail Then Call cmdSendMail(" ")
            cmdSelect_Update_Reset
        End If
    End If

End Sub


Private Sub cmdParam_Update_Ok_Click()
Dim V
App_Debug = "cmdParam_Update_Ok"
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement " & App_Debug): DoEvents

If IsNull(fraParam_Update_Control) Then cmdParam_Update_Ok_Transaction

Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement" & App_Debug): DoEvents

Exit_Sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Update_Quit_Click()
fraParam_Update.Visible = False

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
fraSelect_Update.Visible = False
cmdSelect_Ok.Caption = "Extraire les mouvements"

libRéférenceInterne = ""
If cboSelect_SQL.ListCount > 0 Then
    If currentROPDOSISRV_Rôle = "R" Or currentROPDOSISRV_Rôle = "D" Then
        cboSelect_SQL.ListIndex = 1: lstSelect_Load_1
    Else
        cboSelect_SQL.ListIndex = 0
    End If
End If
'lstSelect_Load_1
If blnOff_Line Then
    YBIATAB0_DATE_CPT_JS1 = "20071030"
    YBIATAB0_DATE_CPT_J = "20071030"
End If
'chkSelect_ROPDOSGECH.Value = "1"
'Call DTPicker_Set(txtSelect_ROPDOSGECH_Max, DateAdd_AMJ("m", 1, YBIATAB0_DATE_CPT_J))
tvwSelect.Visible = False  'True
'cmdSelect_Ok_Click

blnmSelect_Node = False
'''ROPDOS_Aut.Xspécial = False
False_Aut.Avis = False
False_Aut.Comptabiliser = False
False_Aut.Consulter = False
False_Aut.Rapprocher = False
False_Aut.Saisir = False
False_Aut.Swift = False
False_Aut.Valider = False
False_Aut.Virement = False
False_Aut.Xspécial = ROPDOS_Aut.Xspécial


libUpdate_ROPDOSID.ForeColor = &H606060
libUpdate_ROPDOSUUSR.ForeColor = &H606060
'libUpdate_ROPINFID.ForeColor = &H606060

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
htmlFontColor_Blue = "<Font color = #0000FF>"
htmlFontColor_Green = "<Font color = #008080>"
htmlFontColor_Gray = "<Font color = #808080>"
htmlFontColor_Red = "<Font color =#FF0000>"

tvwSelect.Width = 14500
SSTab1.BackColor = &HC0F0FF
blnControl = False
lstW.Visible = False
vDsys = dateImp(DSys)
fraSelect.Enabled = ROPDOS_Aut.Consulter
fraSelect_Update.Top = 0
fraSelect_Update.Left = fraSelect.Width - fraSelect_Update.Width
fraSelect_Update_B.Visible = False
fraSelect_Update_Left = fraTab0.Width - fraSelect_Update.Width - 200
fraSelect_Update_Right = tvwSelect.Width + 200
fraSelect_Update.Left = fraSelect_Update_Left
fraSelect_Update_B.Top = tvwSelect.Top: fraSelect_Update_B.Left = tvwSelect.Left
fraDétail_Update_J.Top = tvwSelect.Top: fraDétail_Update_J.Left = tvwSelect.Left
cmdSelect_Ok.Visible = False
cmdSelect_Update_Close.Visible = False

txtSelect_ROPDOSGECH_Max.Visible = False
chkSelect_ROPDOSGECH.Value = "0"
Call DTPicker_Set(txtSelect_ROPDOSGECH_Max, DateAdd_AMJ("d", 7, DSys))
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
txtSelect_ROPDOSGPRV.AddItem "U - " & usrName_UCase
txtSelect_ROPDOSGPRV.AddItem "V - " & currentROPDOSISRV_Nom
txtSelect_ROPDOSGPRV.AddItem "W - public"

cboSelect_SQL.Clear
If ROPDOS_Aut.Consulter Then
    cboSelect_SQL.AddItem "1  - Extraire"   '" & usrName_UCase & "' + '" & currentROPDOSISRV_Nom & "'"
    'cboSelect_SQL.AddItem "1X - Rechercher un texte"
End If
If ROPDOS_Aut.Saisir Then
    cboSelect_SQL.AddItem "0  - Saisie fiche RO"
    cboSelect_SQL.AddItem "2  - Ouverture d'un dossier"
End If
If ROPDOS_Aut.Xspécial Then
    cboSelect_SQL.AddItem "1M - Liste des modèles"
    cboSelect_SQL.AddItem "2M - Création d'un modèle"
    'cboSelect_SQL.AddItem "2R - Reprise d'un dossier"
    cboSelect_SQL.AddItem "6  - Màj statut des dossiers"
    cboSelect_SQL.AddItem "7  - Etat de suivi des dossiers"
    cboSelect_SQL.AddItem "9$ - spécial JPL : restucturation fichiers "
End If

'_____________________________________________________________________________
sqlYBIATAB0_cboID "ROPDOSSTA", txtSelect_ROPDOSSTA
txtSelect_ROPDOSSTA.AddItem "* - tous"
sqlYBIATAB0_cboID "ROPDOSSTA", txtUpdate_ROPDOSSTA
sqlYBIATAB0_cboID "ROPDOSXDOM", txtUpdate_ROPDOSXDOM
txtUpdate_ROPDOSXDOM.AddItem "?"
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

txtUpdate_ROPINFSTA.Enabled = False
txtUpdate_ROPDOSSTA.Enabled = False
txtUpdate_ROPDOSGTXT.Locked = True
'_____________________________________________________________________________

paramROPDOS_Path = paramServer("\\ROPDOS\") & paramEnvironnement & "\"
paramROPDOS_Path_DROPI = paramServer("\\ROPDOS_DROPI\" & paramEnvironnement & "\")

'_____________________________________________________________________________
lstParam_ROPINFGTXT_Load
lstParam_ROPDOSXDOM_Load
rsYROPINF0_Init zYROPINF0
cmdReset
'_____________________________________________________________________________
txtUpdate_ROPDOSGTXT.Visible = False: cmdSelect_Update_G.Visible = False
txtUpdate_ROPDOSGTXT.Top = fraDétail_Update_B.Top: txtUpdate_ROPDOSGTXT.Left = fraDétail_Update_B.Left
txtUpdate_ROPDOSGTXT.Height = fraDétail_Update_B.Height: txtUpdate_ROPDOSGTXT.Width = fraDétail_Update_B.Width  '- txtUpdate_ROPDOSGTXT.Left * 2
txtUpdate_ROPDOSGTXT.BackColor = &HC0F0FF   ' &HF2F2F2
'Set cmdSelect_Update_G.Container = fraTab0 'txtUpdate_ROPDOSGTXT.Container
cmdSelect_Update_G.Top = 730
cmdSelect_Update_G.Left = 10500
cmdSelect_Update_G.Visible = False
cmdSelect_Update_G.BackColor = &HC0FFC0
libUpdate_ROPINFGTXT.Visible = False
libUpdate_ROPINFGTXT.Top = tvwSelect.Top: libUpdate_ROPINFGTXT.Left = tvwSelect.Left
libUpdate_ROPINFGTXT.Height = tvwSelect.Height ': libUpdate_ROPINFGTXT.Width = tvwSelect.Width
libUpdate_ROPINFGTXT.ForeColor = vbBlue
lstUpdate_ROPINFMAIL.Visible = False
lstUpdate_ROPINFMAIL.Top = tvwSelect.Top: lstUpdate_ROPINFMAIL.Left = tvwSelect.Left
lstUpdate_ROPINFMAIL.Height = tvwSelect.Height ': lstUpdate_ROPINFMAIL.Width = tvwSelect.Width
lstUpdate_ROPINFMAIL.ForeColor = vbBlue
'libUpdate_ROPINFGTXT.BackColor = &HE0E0E0
txtUpdate_ROPINFGTXT_0.ForeColor = vbMagenta
txtUpdate_ROPINFGTXT_0.BackColor = &HC0FFC0
lblUpdate_ROPINFIDTL.ForeColor = vbRed
fraDétail_Update_B.ForeColor = vbMagenta
libUpdate_ROPDOSID.ForeColor = vbRed
libUpdate_ROPDOSIUSR.ForeColor = vbMagenta
libUpdate_ROPDOSGUSR.ForeColor = vbBlue
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
txtUpdate_ROPDOSQUAL.Visible = blnROPDOSQUAL
lblUpdate_ROPDOSQUAL.Visible = blnROPDOSQUAL
If cboSelect_SQL.ListIndex = 0 Then cmdSelect_Ok_Click
End Sub

Private Sub cmdSelect_Update_Close_Click()
Dim V
App_Debug = "cmdselect_Update_Close"
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement " & App_Debug): DoEvents

blnSendMail = False

Select Case cmdSelect_Update_K
    Case "04": cmdSelect_Update_Ok_04:
               chkSelect_Update_B = "0"
               cmdSelect_Update.Clear
               cmdSelect_Update.AddItem " >05 - Ajouter une pièce jointe"
               cmdSelect_Update.ListIndex = 0

    Case "12":
        cmdDétail_Update_Ok
        mROPINFSTA_Value = "F": mROPINFSTA_Set = "+": mROPINFSTA_Where = " ": mROPINFSTAK_Set = "V"
        blnSendMail = True
        If oldYROPINF0.ROPINFGNAT <> "F" Then
            cmdSelect_Update_K = "22"
            cmdDétail_Update_Ok
        Else
            cmdSelect_Update_K = "23"
            cmdSelect_Update_Init_23_Ok
        End If
    Case "13":
        cmdDétail_Update_Ok
        blnSendMail = True
        cmdSelect_Update_K = "23"
        cmdSelect_Update_Init_23_Ok
End Select

libUpdate_ROPINFGTXT.Visible = False
lstUpdate_ROPINFMAIL.Visible = False

Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement" & App_Debug): DoEvents

Exit_Sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Update_G_Click()
Dim blnReset As Boolean
cmdSelect_Update_K = Mid$(cmdSelect_Update_G, 3, 2)
chkSelect_Update_B.Value = "0"
cmdSelect_Update.Clear
cmdSelect_Update.AddItem cmdSelect_Update_G
cmdSelect_Update.Height = 255
blnReset = False
cmdSelect_Update_G.Visible = False
Select Case cmdSelect_Update_K
    Case "03": cmdSelect_Update_Init_03
    Case "14": cmdSelect_Update_Init_14
    Case "24": cmdSelect_Update_Init_24
    Case "64": cmdSelect_Update_Init_64
    Case "98": cmdSelect_Update_Init_98
    Case "99": cmdSelect_Update_Init_99: blnReset = True
End Select

If blnReset Then fraSelect_Update.Visible = False

End Sub

Private Sub cmdSelect_Update_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If blnSelect_Update_Expandable Then
    If cmdSelect_Update.ListCount > 1 And cmdSelect_Update.Height < 400 Then cmdSelect_Update.Height = 250 * cmdSelect_Update.ListCount
End If
End Sub

Private Sub dirListBox_Change()
filDoc.PATH = dirListBox.PATH
filDoc.Pattern = "*.*"
End Sub


Public Sub lstParam_ROPDOSXAPP_Load(lK1 As String)
Dim X As String, X12 As String
lstParam_ROPDOSXAPP.Clear
X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID =  'ROPDOSXAPP'" _
    & " and BIATABK1 = '" & lK1 & "'"
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    X12 = rsSab("BIATABK2")
    lstParam_ROPDOSXAPP.AddItem X12 & " - " & Trim(Mid$(rsSab("BIATABTXT"), 1, 24))
    rsSab.MoveNext
Loop

If lstParam_ROPDOSXAPP.ListCount > 0 Then lstParam_ROPDOSXAPP.ListIndex = 0
End Sub

Public Sub lstParam_ROPDOSXDOM_Load()
Dim X As String, X12 As String
lstParam_ROPDOSXDOM.Clear
X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID =  'ROPDOSXDOM'"
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    X12 = rsSab("BIATABK1")
    lstParam_ROPDOSXDOM.AddItem X12 & " - " & Trim(Mid$(rsSab("BIATABTXT"), 1, 24))
    rsSab.MoveNext
Loop

If lstParam_ROPDOSXDOM.ListCount > 0 Then lstParam_ROPDOSXDOM.ListIndex = 0
End Sub

Public Sub lstParam_ROPINFGTXT_Load()
Dim X As String, X12 As String * 12
Dim mBIATABK1 As String

mBIATABK1 = ""
lstParam_ROPINFGTXT.Clear
libUpdate_ROPINFGTXT.Clear

X = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID =  'ROPINFGTXT' "
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
        libUpdate_ROPINFGTXT.AddItem "___________ " & UCase(Trim(X12)) & " ___________________________________"
    End If
    libUpdate_ROPINFGTXT.AddItem Trim(rsSab("BIATABTXT"))
    rsSab.MoveNext
Loop

'If lstParam_ROPINFGTXT.ListCount > 0 Then lstParam_ROPINFGTXT.ListIndex = 0
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
blnOk = Not tvwSelect.Visible
Call lstErr_Clear(lstErr, cmdContext, "> SAB_CDR_cmdSelect_Ok ........"): DoEvents
cmdSelect_Ok.Visible = False

fraSelect_Update.Visible = False
fraSelect_Update_B.Visible = False
libUpdate_ROPINFGTXT.Visible = False
lstUpdate_ROPINFMAIL.Visible = False
fraSelect_Options_1.Enabled = False
lstUpdate_Modèle.Visible = False

'tvwSelect.Nodes.Clear
tvwSelect.Visible = False

blnDossierModèle = False
blnDossierReprise = False
DoEvents
If blnOk Then
    cmdSelect_Ok.Caption = "Modifier les options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options_1.BackColor = &HE0FFFF      '&H8000000F
    Call usrColor_Container(fraSelect_Options_1, fraSelect_Options_1.BackColor)
    cmdSelect_SQL_X1 = Mid$(cboSelect_SQL, 1, 1)
    cmdSelect_SQL_K = Trim(Mid$(cboSelect_SQL, 1, 2))
    Select Case cmdSelect_SQL_K
        Case "0": cmdSelect_SQL_0
        Case "1": Call cmdSelect_SQL_1(0)
        'Case "1U": Call cmdSelect_SQL_1(0)
        'Case "1V": Call cmdSelect_SQL_1(0)
        'Case "1W": Call cmdSelect_SQL_1(0)
        'Case "1X": Call cmdSelect_SQL_1X
        Case "1M":    blnDossierModèle = True: Call cmdSelect_SQL_1(0)
        Case "2": Call cmdSelect_SQL_1(0)
        Case "2M": cmdSelect_SQL_2M
        'Case "2R": cmdSelect_SQL_2R
        Case "6": cmdSelect_SQL_6
        Case "7": cmdSelect_SQL_7
        Case "9$": cmdSelect_SQL_JPL

    End Select

    tvwSelect.Enabled = True
Else
    cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
    cmdSelect_Ok.BackColor = &HC0F0FF    '&HE0FFFF
    fraSelect_Options_1.BackColor = &HC0F0FF   ' &HE0FFFF  ' &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options_1, fraSelect_Options_1.BackColor)
    tvwSelect.Visible = False
    tvwSelect.Enabled = False
    fraSelect_Options_1.Enabled = True

End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_CDR_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
cmdSelect_Ok.Visible = True

End Sub


Private Sub cmdSelect_SQL_1(lROPDOSID As Long)
Dim V, X As String
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String, blnAnd As Boolean, filtreROPDOSGUSR As String, filtreROPINFGUSR As String
Dim xROPDOSID As String
Dim blnROPDOSGUSR As Boolean, wId As Long
Dim blnFiltre As Boolean
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_1"): DoEvents

currentAction = "cmdSelect_SQL_1"
xWhere = "": xAnd = "": filtreROPDOSGUSR = ""
blnROPDOSGUSR = False
blnFiltre = False
If lROPDOSID > 0 Then
    xROPDOSID = lROPDOSID
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
            xROPDOSID = Trim(txtSelect_ROPDOSID)
            If xROPDOSID <> "" Then xWhere = " and   ROPDOSId =" & xROPDOSID
        Case Else: blnFiltre = True
    End Select
End If

If blnFiltre Then
Select Case Mid$(txtSelect_ROPDOSGPRV, 1, 1)
    Case "U": cmdSelect_SQL_K = "1U"
    Case "V": cmdSelect_SQL_K = "1V"
    Case "W": cmdSelect_SQL_K = "1W"
End Select
    xROPDOSID = Trim(txtSelect_ROPDOSID)
    If xROPDOSID <> "" And Val(xROPDOSID) >= 1000 Then
        xWhere = " Where   ROPDOSId =" & xROPDOSID
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
    
    Call DTPicker_Control(txtSelect_ROPDOSGECH_Max, wAmjMax)
    
    If chkSelect_ROPDOSGECH = "1" Then
        xWhere = xWhere & " and   ROPDOSgech <= '" & wAmjMax & "'"
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
    Case "1": filtreROPDOSGUSR = " and ( ROPDOSIUSR = '" & usrName_UCase & "' or ROPDOSGUSR = '" & usrName_UCase & "'" _
                   & " or ROPDOSISRV = '" & currentROPDOSISRV & "' or ROPDOSGSRV = '" & currentROPDOSISRV & " ' or   ROPDOSGPRV = 'W')"
               filtreROPINFGUSR = " and ROPINFGSRV = '" & currentROPDOSISRV & "'"
               blnROPDOSGUSR = True
    Case "1U": filtreROPDOSGUSR = " and ROPDOSGPRV = 'U' and  ROPDOSIUSR = '" & usrName_UCase & "'"
               filtreROPINFGUSR = "" '" and ROPINFGUSR = '" & usrName_UCase & "'"
               blnROPDOSGUSR = False
    Case "1V": filtreROPDOSGUSR = " and ROPDOSGPRV = 'V' and ( ROPDOSISRV = '" & currentROPDOSISRV & "' or ROPDOSGSRV = '" & currentROPDOSISRV & "')"
               filtreROPINFGUSR = "" '" and ROPINFGSRV = '" & currentROPDOSISRV & "'"
               blnROPDOSGUSR = False
    Case "1W": filtreROPDOSGUSR = " and  ROPDOSGPRV = 'W'"
End Select
'____________________________________________________________________________________
cmdSelect_SQL_1_Ok:

If xWhere <> "" Then
    Mid$(xWhere, 1, 6) = " where"
Else
    If xAnd <> "" Then Mid$(xAnd, 1, 6) = " where"
End If

xSql = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 " & xWhere & xAnd & filtreROPDOSGUSR
Set rsSab = cnsab.Execute(xSql)
ReDim selYROPDOS0(100): selYROPDOS0_Nb = 0
Do While Not rsSab.EOF
        If selYROPDOS0_Nb >= UBound(selYROPDOS0) Then ReDim Preserve selYROPDOS0(selYROPDOS0_Nb + 100)
        selYROPDOS0_Nb = selYROPDOS0_Nb + 1
        V = rsYROPDOS0_GetBuffer(rsSab, selYROPDOS0(selYROPDOS0_Nb))
    rsSab.MoveNext
Loop
'____________________________________________________________________________________
If blnROPDOSGUSR Then
    xWhere = Replace(xWhere, "DOS", "INF")
   ''''''''''''''''''''''''''''' xwhere = xwhere & " and ropinfgtxt like '%'
    xSql = "select  distinct(ROPINFID) from " & paramIBM_Library_SABSPE & ".YROPINF0 " _
        & xWhere & filtreROPINFGUSR
    X = ""
    Set rsSab = cnsab.Execute(xSql)
    Do While Not rsSab.EOF
        wId = rsSab("ROPINFID")
        For K = 1 To selYROPDOS0_Nb
           If wId = selYROPDOS0(K).ROPDOSID Then wId = 0: Exit For
        Next K
        If wId > 0 Then X = X & wId & ","
        rsSab.MoveNext
    Loop
    If X <> "" Then
        xWhere = "Where ROPDOSID in(" & Mid$(X, 1, Len(X) - 1) & ")"
       xSql = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 " & xWhere & xAnd
       Set rsSab = cnsab.Execute(xSql)
       Do While Not rsSab.EOF
            If selYROPDOS0_Nb >= UBound(selYROPDOS0) Then ReDim Preserve selYROPDOS0(selYROPDOS0_Nb + 100)
            selYROPDOS0_Nb = selYROPDOS0_Nb + 1
            V = rsYROPDOS0_GetBuffer(rsSab, selYROPDOS0(selYROPDOS0_Nb))
           rsSab.MoveNext
       Loop
    End If
End If
'____________________________________________________________________________________

cmdSelect_SQL_1_YROPDOS0
'____________________________________________________________________________________

If xROPDOSID <> "" Then
    X = "D" & Format$(CLng(xROPDOSID), "000000000")
    Set xNode = tvwSelect.Nodes(X)
    tvwSelect_NodeClick xNode
    X = "D" & Format$(CLng(xROPDOSID), "000000000") & "000010000000001"
    Set xNode = tvwSelect.Nodes(X)
    '$20071122_JPL$ tvwSelect_NodeClick xNode
    fraSelect_Update.Visible = False
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub cmdSelect_SQL_1X()
Dim V, X As String
Dim xSql As String, K As Long
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
xSql = "select  distinct(ROPINFID) from " & paramIBM_Library_SABSPE & ".YROPINF0 " _
    & xWhere
Set rsSab = cnsab.Execute(xSql)
X = ""
Do While Not rsSab.EOF
    wId = rsSab("ROPINFID")
    If wId > 0 Then X = X & wId & ","
    rsSab.MoveNext
Loop
ReDim selYROPDOS0(100): selYROPDOS0_Nb = 0

If X <> "" Then
    xWhere = "Where ROPDOSID in(" & Mid$(X, 1, Len(X) - 1) & ")"
   xSql = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 " & xWhere & xAnd
   Set rsSab = cnsab.Execute(xSql)
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

tvwSelect.Visible = False
blnSelect_Update_B_Display = True
Me.Enabled = True: Me.MousePointer = 0
fraSelect_Update.Left = fraSelect_Update_Left

cmdSelect_Update_K = "04"

fraSelect_Display
fraDétail_Display
cmdSelect_Update_Init_04

fraDétail_Update_B.Enabled = True
cmdSelect_Update.Visible = False
cmdSelect_Update_Ok.Visible = True
cmdSelect_Update_Close.Caption = "Enregistrer + ajouter une pièce jointe"
cmdSelect_Update_Close.Visible = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_Update_Click()
Dim xItem As String, K As Integer
cmdSelect_Update_Ok.Visible = False
cmdSelect_Update_Close.Visible = False
blnYROPINF0_12X = False

xItem = Mid$(cmdSelect_Update, 3, Len(cmdSelect_Update))
cmdSelect_Update_K = Mid$(cmdSelect_Update, 3, 2)
If blnControl Then
    fraDétail_Update_J.Visible = False
    fraSelect_Update_B.Visible = False
    blnSelect_Update_Expandable = False
    blnSendMail = True 'False
    If cmdSelect_Update_K = "00" Then
        tvwSelect_Display
        
    Else
        If blnSelect_Update_EnCours Then tvwSelect_Display
        blnSelect_Update_EnCours = True
        Select Case cmdSelect_Update_K
            Case "01": cmdSelect_Update_Init_01
            Case "02":
                    Select Case Mid$(cmdSelect_Update, 5, 1)
                        Case "F": Call cmdSelect_Update_Init_02("F")
                        Case "I": blnROPINFIDT_Insérer = True: cmdSelect_Update_Init_02 ("A")
                        Case Else:  Call cmdSelect_Update_Init_02("A")
                    End Select
                    
            Case "03": cmdSelect_Update_Init_03
            Case "05": cmdSelect_Update_Init_05
            Case "11": cmdSelect_Update_Init_11
            Case "12":
                        If Mid$(cmdSelect_Update, 5, 1) = "X" Then
                            blnYROPINF0_12X = True
                            cmdSelect_Update_Init_11
                        Else
                            cmdSelect_Update_Init_11
                            If oldYROPINF0.ROPINFGNAT = "F" Then
                                 cmdSelect_Update_Close.Caption = "Enregistrer + Clôturer ce processus"
                            Else
                                 cmdSelect_Update_Close.Caption = "Enregistrer + Clôturer cette action"
                            End If
                            If Not blnDossierModèle Then cmdSelect_Update_Close.Visible = True
                       End If
            Case "13": cmdSelect_Update_Init_11
                       cmdSelect_Update_Close.Caption = "Enregistrer + Clôturer ce processus"
                       cmdSelect_Update_Close.Visible = True
            Case "14": cmdSelect_Update_Init_14
            Case "21": cmdSelect_Update_Init_21
            Case "22": cmdSelect_Update_Init_22
            Case "23": cmdSelect_Update_Init_23
            Case "24": cmdSelect_Update_Init_24
            Case "31": cmdSelect_Update_Init_31
            Case "32": cmdSelect_Update_Init_32
            Case "33": cmdSelect_Update_Init_33
            Case "34": cmdSelect_Update_Init_34
            Case "41": cmdSelect_Update_Init_41
            Case "42": cmdSelect_Update_Init_42
            Case "43": cmdSelect_Update_Init_43
            Case "44": cmdSelect_Update_Init_44
            Case "51": cmdSelect_Update_Init_51
            Case "52": cmdSelect_Update_Init_52
            Case "53": cmdSelect_Update_Init_53
            Case "54": cmdSelect_Update_Init_54
            Case "64": cmdSelect_Update_Init_64
            Case "74": cmdSelect_Update_Init_74
            Case "98": cmdSelect_Update_Init_98
            Case "99": cmdSelect_Update_Init_99
            Case "00", "10", "20", "30", "40", "50", "60":
            Case Else: MsgBox " NON Géré"
        End Select
    End If
    Call cmdSelect_Update_Display(xItem)
    If fraSelect_Update.Visible Then
        If txtUpdate_ROPINFGTXT.Enabled Then
            txtUpdate_ROPINFGTXT.SetFocus
        Else
            cmdSelect_Update_Quit.SetFocus
        End If
    End If
    cmdSelect_Update.Height = 280
End If
End Sub


Private Sub cmdSelect_Update_Ok_Click()
Dim V
App_Debug = "cmdselect_Update_Ok"
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement " & App_Debug): DoEvents

Select Case cmdSelect_Update_K
    Case "01", "02", "03", "05": cmdDétail_Update_Ok_Insert
    Case "04": cmdSelect_Update_Ok_04
    Case "11", "12", "13": cmdDétail_Update_Ok
    Case "14": cmdSelect_Update_Ok_14
    Case "74": cmdSelect_Update_Ok_Transaction_Duplication
    Case "98": cmdSelect_Update_Ok_98
End Select
libUpdate_ROPINFGTXT.Visible = False
lstUpdate_ROPINFMAIL.Visible = False

Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement" & App_Debug): DoEvents

Exit_Sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Update_Quit_Click()

fraDétail_Update_J.Visible = False
fraSelect_Update_B.Visible = False
libUpdate_ROPINFGTXT.Visible = False
lstUpdate_ROPINFMAIL.Visible = False
lstUpdate_Modèle.Visible = False
cmdSelect_Update.Height = 250 * cmdSelect_Update.ListCount

If Not cmdSelect_Update_Ok.Visible Then
    blntvwSelect_Click = False
        cmdSelect_Update_Reset

    Exit Sub
End If
cmdSelect_Update_Ok.Visible = False
cmdSelect_Update_Close.Visible = False

If blnSelect_Update_EnCours Then
    cmdSelect_Update_Reset
Else
    fraSelect_Update.Visible = False
End If

End Sub


Private Sub cmdSelect_Update_Reset()
fraDétail_Update_J.Visible = False
fraSelect_Update_B.Visible = False
cmdSelect_Update_Ok.Visible = False
cmdSelect_Update_Close.Visible = False
If blnSelect_Update_EnCours Then tvwSelect_Display
blnSelect_Update_Expandable = True
cmdSelect_Update.Height = 255 * cmdSelect_Update.ListCount
blnSelect_Update_B_Display = False
'$20071122_JPL$ chkSelect_Update_B.Value = "1"
fraSelect_Update.Visible = False
End Sub

Private Sub DriveListBox_Change()
On Error Resume Next
dirListBox.PATH = DriveListBox.Drive ' .PATH
End Sub

Private Sub filDoc_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
oldFileName = filDoc.PATH & "\" & filDoc.FileName
newDirPath = paramROPDOS_Path & oldYROPINF0.ROPINFID
newFileName = filDoc.FileName
newFileExtension = fileName_Extension(filDoc.FileName)
'txtUpdate_ROPINFGUSR = newFileExtension
txtUpdate_ROPINFGTXT = filDoc.FileName
fraDétail_Update_J.Visible = False
txtUpdate_ROPINFGTXT.Locked = False
fraSelect_Update.Left = fraSelect_Update_Right
blntvwSelect_Click = True
Me.Enabled = True: Me.MousePointer = 0
On Error Resume Next
cmdSelect_Update_Ok.SetFocus
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
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub


Private Sub fraSelect_Update_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If fraSelect_Update.Left > 9000 Then fraSelect_Update.Left = fraSelect_Update_Left
'blntvwSelect_Click = False
End Sub

Private Sub libUpdate_ROPINFGTXT_Click()
On Error Resume Next
txtUpdate_ROPINFGTXT = txtUpdate_ROPINFGTXT & libUpdate_ROPINFGTXT
txtUpdate_ROPINFGTXT.SetFocus
txtUpdate_ROPINFGTXT.SelStart = Len(txtUpdate_ROPINFGTXT)
libUpdate_ROPINFGTXT.Visible = False
End Sub



Private Sub lstAut_ROPDOSISRV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim K As Integer, K2 As Integer

Me.Enabled = False: Me.MousePointer = vbHourglass
lstAut_ROPDOSISRV_Display
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub lstAut_Usr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim K As Integer, K2 As Integer
Dim wIndex As Long, xMail As String

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

Call lstAut_ROPDOSGUSR_Display(wIndex)
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


Private Sub lstParam_ROPDOSXAPP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
fraParam_Reset
oldParam.BIATABID = "ROPDOSXAPP"
oldParam.BIATABK1 = Mid$(lstParam_ROPDOSXDOM.Text, 1, 12)
oldParam.BIATABK2 = Mid$(lstParam_ROPDOSXAPP.Text, 1, 12)
blnParam_ROPDOSXDOM = False
If lstParam_ROPDOSXAPP.ListCount > 0 Then
    mnuParam_Delete.Enabled = True
Else
    mnuParam_Delete.Enabled = False
End If
lstParam_ROPDOSXAPP.Enabled = True
Me.PopupMenu mnuParam, vbPopupMenuLeftButton

End Sub


Private Sub lstParam_ROPDOSXDOM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
fraParam_Reset
oldParam.BIATABID = "ROPDOSXDOM"
oldParam.BIATABK1 = Mid$(lstParam_ROPDOSXDOM.Text, 1, 12)

blnParam_ROPDOSXDOM = True
lstParam_ROPDOSXAPP_Load (oldParam.BIATABK1)
lstParam_ROPDOSXAPP.Enabled = True
If lstParam_ROPDOSXAPP.ListCount > 0 Then
    mnuParam_Delete.Enabled = False
Else
    mnuParam_Delete.Enabled = True
End If
Me.PopupMenu mnuParam, vbPopupMenuLeftButton

End Sub


Private Sub lstParam_ROPINFGTXT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
fraParam_Reset
blnParam_ROPINFGTXT = True
oldParam.BIATABID = "ROPINFGTXT"
If lstParam_ROPINFGTXT.ListCount > 0 And lstParam_ROPINFGTXT.ListIndex >= 0 Then
    oldParam.BIATABK1 = Mid$(lstParam_ROPINFGTXT.Text, 1, 12)
 '   oldParam.BIATABK2 = arrROPINFGTXT_BIATABK2(lstParam_ROPINFGTXT.ListIndex)
    mnuParam_Delete.Enabled = True
    Me.PopupMenu mnuParam, vbPopupMenuLeftButton
Else
    mnuParam_Insert_Click
    
End If
End Sub

Private Sub lstUpdate_Modèle_Click()
lstUpdate_Modèle.Visible = False
cmdSelect_Update_Ok.Visible = True
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
Call BiaPgmAut_Init(mAPP_Menu, ROPDOS_Aut)

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
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim wDétail As String
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
    lstUpdate_ROPINFMAIL.Visible = False
    mailYROPINF0.ROPINFIDP = -1
    
    For K = 0 To lstUpdate_ROPINFMAIL.ListCount - 1
        lstUpdate_ROPINFMAIL.ListIndex = K
        If lstUpdate_ROPINFMAIL.Selected(K) Then Call cmdSendMail_Recipient(Trim(lstUpdate_ROPINFMAIL.Text))
    Next K
    lstUpdate_ROPINFMAIL.Visible = True
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

'=============================================

If wRecipient = "" Then Exit Sub
'_____________________________________________________________________________________________
Call arrYROPINF0_SQL(mailYROPDOS0.ROPDOSID)
'_____________________________________________________________________________________________
wSendMail.From = currentZMNUUTI0.MNUUTIMAI
wSendMail.FromDisplayName = usrName_UCase
wSendMail.Recipient = wRecipient
wSendMail.CcRecipient = wccRecipient
wSendMail.Attachment = ""

bgColor = "" '"cyan"

'wSendMail.Subject = "RO -" & mailYROPINF0.ROPINFID & "-" & mailYROPINF0.ROPINFIDP & "-" & mailYROPINF0.ROPINFIDT & "-" & mailYROPINF0.ROPINFIDT2 & " : " & mailSubject
wSendMail.Subject = "BIA.RO -" & mailYROPDOS0.ROPDOSID & " : " & Trim(mailYROPDOS0.ROPDOSXDOM) & " - " & Trim(mailYROPDOS0.ROPDOSXAPP) & " - " & Trim(mailYROPDOS0.ROPDOSXID) _
                   & " (" & Trim(tvwSelect_Display_USR(mailYROPDOS0.ROPDOSIUSR)) & ")"
'wDétail = "<FONT color=#0000A0  face=" & Asc34 & "Arial" & Asc34 & ">" _
'_____________________________________________________________________________________________
mFontFace = "<FONT face=" & Asc34 & "Arial Unicode MS" & Asc34 & ">"
'wDétail = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">" _
'        & mFontFace _
wDétail = "<TABLE border = 1  width=900 height=5 bgcolor=#0000FF cellpadding=4 ><TR>" _
         & "<TD  width=100 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Référence</TD>" _
         & "<TD  width=600 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Responsable</B></TD>" _
         & "<TD  width=100 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Echéance</TD>" _
         & "<TD  width=100 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Statut</TD>" _
        & "</TR></TABLE>"

wTD_BackColor = "bgcolor = #87CEFA"
wTD_Sta_ForeColor = wTD_BackColor
If Trim(mailYROPDOS0.ROPDOSXID) = "" Then
    wROPDOSXID = ""
Else
    wROPDOSXID = " . " & htmlFontColor_Red & mailYROPDOS0.ROPDOSXID
End If

X = cmdSendMail_Sta(mailYROPDOS0.ROPDOSSTA, wTD_Sta_ForeColor)
xUsr = Trim(cmdSendMail_USR(mailYROPDOS0.ROPDOSGUSR)) & htmlFontColor_Red & " . . . " & mailYROPDOS0.ROPDOSXDOM & " - " & mailYROPDOS0.ROPDOSXAPP
wDétail = wDétail & "<TABLE border = 0  width=900   cellpadding=4 ><Font  color=#FFFFFF><TR>" _
         & "<TD " & wTD_BackColor & " width=100 height=5><span style='font-size:10.0pt;font-family:Arial'>" & htmlFontColor_Blue & mailYROPDOS0.ROPDOSID & wROPDOSXID & "</TD>" _
         & "<TD " & wTD_BackColor & " width=600 height=5><span style='font-size:10.0pt;font-family:Arial'>" & htmlFontColor_Blue & "<B>" & xUsr & "</B></TD>" _
         & "<TD " & wTD_BackColor & " width=100 height=5><span style='font-size:10.0pt;font-family:Arial'>" & cmdSendMail_Ech(mailYROPDOS0.ROPDOSGECH, mailYROPDOS0.ROPDOSSTA) & "</TD>" _
         & "<TD " & wTD_Sta_ForeColor & " width=100 height=5><span style='font-size:10.0pt;font-family:Arial'>" & X & "</TD>" _
        & "</TR></TABLE>"
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
            wTD_Txt_ForeColor = htmlFontColor_Blue
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
                wTD_BackColor = "bgcolor = #90FFFF"
                wTD_Sta_ForeColor = wTD_BackColor
                X = cmdSendMail_Sta(xYROPINF0.ROPINFSTA, wTD_Sta_ForeColor)
                xUsr = Trim(cmdSendMail_USR(xYROPINF0.ROPINFGUSR))
                wDétail = wDétail & "<NOBR><TABLE  width=900 border=1   cellpadding=4 ></B><TR>" _
                         & "<TD " & wTD_BackColor & " width=100 height=5><span style='font-size:8.0pt;font-family:Arial'>" & wTD_ForeColor & libROPINFGNAT & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=600 height=5><span style='font-size:8.0pt;font-family:Arial'>" & wTD_ForeColor & xUsr & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=100 height=5><span style='font-size:8.0pt;font-family:Arial'>" & cmdSendMail_Ech(xYROPINF0.ROPINFGECH, xYROPINF0.ROPINFSTA) & "</TD>" _
                         & "<TD " & wTD_Sta_ForeColor & " width=100 height=5><span style='font-size:8.0pt;font-family:Arial'>" & X & "</TD>" _
                         & "</TR>" _
                         & "<TR>" _
                         & "<TD colspan=4  height=5><PRE><span style='font-size:10.0pt;font-family:Arial'>" & wTD_Txt_ForeColor & cmdSendMail_Txt(xYROPINF0.ROPINFGTXT, "H") _
                         & "</TD></TR></TABLE>"
             Case "A", "F":
                libROPINFGNAT = xYROPINF0.ROPINFID & " § " & xYROPINF0.ROPINFIDP & " - " & Format$(xYROPINF0.ROPINFIDT, "00")
             
                wTD_BackColor = "bgcolor =  #B0FFFF"
                wTD_Sta_ForeColor = wTD_BackColor
                X = cmdSendMail_Sta(xYROPINF0.ROPINFSTA, wTD_Sta_ForeColor)
                'If xYROPINF0.ROPINFIDT = mailYROPINF0_Suivant.ROPINFIDT And mailYROPINF0_Suivant.ROPINFIDT > 0 Then
                '    wTD_Sta_ForeColor = "bgcolor =#FFB000"
                'End If

                xUsr = cmdSendMail_USR(xYROPINF0.ROPINFGUSR)
                wDétail = wDétail & "<NOBR><TABLE  width=900 border=1    cellpadding=4 ></B><TR>" _
                         & "<TD " & wTD_BackColor & " width=100  height=5><span style='font-size:8.0pt;font-family:Arial'>" & wTD_ForeColor & libROPINFGNAT & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=600  height=5><span style='font-size:8.0pt;font-family:Arial'>" & wTD_ForeColor & xUsr & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=100  height=5><span style='font-size:8.0pt;font-family:Arial'>" & cmdSendMail_Ech(xYROPINF0.ROPINFGECH, xYROPINF0.ROPINFSTA) & "</TD>" _
                         & "<TD " & wTD_Sta_ForeColor & " width=100  height=5><span style='font-size:8.0pt;font-family:Arial'>" & X & "</TD>" _
                         & "</TR>" _
                         & "<TR>" _
                         & "<TD colspan=4  height=5><PRE><span style='font-size:10.0pt;font-family:Arial'>" & wTD_Txt_ForeColor & cmdSendMail_Txt(xYROPINF0.ROPINFGTXT, "H") _
                         & "</TD></TR></TABLE>"
              Case "J"
                    'X = fraDétail_Display_PJ_FileName(xYROPINF0.ROPINFGUSR, False)
                    'wSendMail.Attachment = wSendMail.Attachment & X & ";"
                    wTD_BackColor = "bgcolor = #F5F5F5"
                    wTD_Sta_ForeColor = wTD_BackColor
                    X = cmdSendMail_Sta(xYROPINF0.ROPINFSTA, wTD_Sta_ForeColor)
                    xUsr = cmdSendMail_USR(xYROPINF0.ROPINFGUSR)

                    wPJ = "<a href=" & Asc34 _
                        & paramROPDOS_Path_DROPI & xYROPINF0.ROPINFID _
                        & "\" & xYROPINF0.ROPINFID & "_" & xYROPINF0.ROPINFIDP _
                        & "_" & xYROPINF0.ROPINFIDT & "_" & xYROPINF0.ROPINFIDT2 & "." & UCase$(Trim(xYROPINF0.ROPINFGUSR)) _
                        & Asc34 & ">" & Trim(xYROPINF0.ROPINFGTXT)

                    wDétail = wDétail & "<TABLE   width=900 border=1 cellpadding=4 ></B><TR>" _
                         & "<TD " & wTD_BackColor & " width=100 height=5><span style='font-size:8.0pt;font-family:Arial'>" & wTD_Txt_ForeColor & "PJ" & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=600 height=5><span style='font-size:8.0pt;font-family:Arial'>" & wTD_Txt_ForeColor & wPJ & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=100 height=5><span style='font-size:8.0pt;font-family:Arial'>" & wTD_Txt_ForeColor & dateImp10(xYROPINF0.ROPINFGECH) & "</TD>" _
                         & "<TD " & wTD_Sta_ForeColor & " width=100 height=5><span style='font-size:8.0pt;font-family:Arial'>" & X & "</TD>" _
                         & "</TD></TR></TABLE>"
               Case Else
                    Select Case xYROPINF0.ROPINFGNAT
                        Case "N": libROPINFGNAT = "Note"
                        Case Else: libROPINFGNAT = xYROPINF0.ROPINFGNAT
                    End Select
                    wTD_BackColor = "bgcolor = #F5F5F5"
                    wTD_Sta_ForeColor = wTD_BackColor
                    X = cmdSendMail_Sta(xYROPINF0.ROPINFSTA, wTD_Sta_ForeColor)
                    xUsr = cmdSendMail_USR(xYROPINF0.ROPINFGUSR)
                    wDétail = wDétail & "<TABLE   width=900 border=1 cellpadding=4 ></B><TR>" _
                         & "<TD " & wTD_BackColor & " width=100 height=5><span style='font-size:8.0pt;font-family:Arial'>" & wTD_Txt_ForeColor & libROPINFGNAT & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=600 height=5><span style='font-size:8.0pt;font-family:Arial'>" & wTD_Txt_ForeColor & xUsr & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=100 height=5><span style='font-size:8.0pt;font-family:Arial'>" & wTD_Txt_ForeColor & dateImp10(xYROPINF0.ROPINFGECH) & "</TD>" _
                         & "<TD " & wTD_Sta_ForeColor & " width=100 height=5><span style='font-size:8.0pt;font-family:Arial'>" & X & "</TD>" _
                         & "</TR>" _
                         & "<TR>" _
                         & "<TD colspan=4  height=5><PRE><span style='font-size:10.0pt;font-family:Arial'>" & wTD_Txt_ForeColor & cmdSendMail_Txt(xYROPINF0.ROPINFGTXT, "H") _
                         & "</TD></TR></TABLE>"
               End Select
            
    End If
Next K

wSendMail.Message = wDétail
wSendMail.AsHTML = True
'MsgBox "cmdsensmail Exit"
'Exit Sub
If Not blnOff_Line And mailYROPINF0.ROPINFID > 2000 Then
    srvSendMail.Monitor wSendMail
    For K1 = 1 To arrRecipient_Nb
        lstErr.AddItem "@ " & arrRecipient(K1)
    Next K1

End If

End Sub

Public Sub cmdSendMail_X(lFct As String)
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim wDétail As String
Dim libROPINFGNAT As String
Dim K As Long, K1 As Long, blnDisplay As Boolean
Dim wROPINFSTA As String, wTD_BackColor As String, wSta_BackColor As String, wTxt_BackColor As String
Dim mFontColor_Blue As String, mFontColor_Green As String

Dim X1 As String, wUsrName As String
Dim xUsr As String, kLen As Integer, K2 As Integer
Dim meName_Ucase As String

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
    lstUpdate_ROPINFMAIL.Visible = False
    mailYROPINF0.ROPINFIDP = -1
    
    For K = 0 To lstUpdate_ROPINFMAIL.ListCount - 1
        lstUpdate_ROPINFMAIL.ListIndex = K
        If lstUpdate_ROPINFMAIL.Selected(K) Then Call cmdSendMail_Recipient(Trim(lstUpdate_ROPINFMAIL.Text))
    Next K
    lstUpdate_ROPINFMAIL.Visible = True
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

'=============================================

If wRecipient = "" Then Exit Sub
'_____________________________________________________________________________________________
Call arrYROPINF0_SQL(mailYROPDOS0.ROPDOSID)
'_____________________________________________________________________________________________
wSendMail.From = currentZMNUUTI0.MNUUTIMAI
wSendMail.FromDisplayName = usrName_UCase
wSendMail.Recipient = wRecipient
wSendMail.CcRecipient = wccRecipient
wSendMail.Attachment = ""

bgColor = "" '"cyan"
mFontColor_Blue = "<Font color = #0000FF>"
mFontColor_Green = "<Font color = #008080>"
'wSendMail.Subject = "RO -" & mailYROPINF0.ROPINFID & "-" & mailYROPINF0.ROPINFIDP & "-" & mailYROPINF0.ROPINFIDT & "-" & mailYROPINF0.ROPINFIDT2 & " : " & mailSubject
wSendMail.Subject = "BIA.RO -" & mailYROPDOS0.ROPDOSID & " : " & Trim(mailYROPDOS0.ROPDOSXDOM) & " - " & Trim(mailYROPDOS0.ROPDOSXAPP) & " - " & Trim(mailYROPDOS0.ROPDOSXID) _
                   & " (" & Trim(tvwSelect_Display_USR(mailYROPDOS0.ROPDOSIUSR)) & ")"
'wDétail = "<FONT color=#0000A0  face=" & Asc34 & "Arial" & Asc34 & ">" _
'_____________________________________________________________________________________________
wDétail = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">" _
        & "<FONT face=" & Asc34 & prtFontName_CourierNew & Asc34 & ">" _
        & "<TABLE border = 1  width=800 height= 10 bgcolor=#0000FF cellpadding=5 ><TR>" _
         & "<TD  width=200 height=10><Font color=#FFFFFF>Référence</TD>" _
         & "<TD  width=400 height=10><Font color=#FFFFFF>Responsable</B></TD>" _
         & "<TD  width=100 height=10><Font color=#FFFFFF>Echéance</TD>" _
         & "<TD  width=100 height=10><Font color=#FFFFFF>Statut</TD>" _
        & "</TR></TABLE>"

wTD_BackColor = "bgcolor = #87CEFA"
wSta_BackColor = wTD_BackColor
X = cmdSendMail_Sta(mailYROPDOS0.ROPDOSSTA, wSta_BackColor)
xUsr = Trim(cmdSendMail_USR(mailYROPDOS0.ROPDOSGUSR))
wDétail = wDétail & "<TABLE border = 0  width=800   cellpadding=5 ><Font  color=#FFFFFF><TR>" _
         & "<TD " & wTD_BackColor & " width=200 height=10><Font size=2>" & mFontColor_Blue & "Dossier : " & mailYROPDOS0.ROPDOSID & "</TD>" _
         & "<TD " & wTD_BackColor & " width=400 height=10><Font size=2>" & mFontColor_Blue & "<B>" & xUsr & "</B></TD>" _
         & "<TD " & wTD_BackColor & " width=100 height=10><Font size=2>" & cmdSendMail_Ech(mailYROPDOS0.ROPDOSGECH, mailYROPDOS0.ROPDOSSTA) & "</TD>" _
         & "<TD " & wSta_BackColor & " width=100 height=10><Font size=2>" & X & "</TD>" _
        & "</TR></TABLE>"
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
        libROPINFGNAT = xYROPINF0.ROPINFIDP & " - " & Format$(xYROPINF0.ROPINFIDT, "00")
        If xYROPINF0.ROPINFIDP = mailYROPINF0.ROPINFIDP Then
           ' If xYROPINF0.ROPINFIDT = mailYROPINF0.ROPINFIDT Then
           '     If xYROPINF0.ROPINFIDT2 = mailYROPINF0.ROPINFIDT2 Then wTxt_BackColor = "<font color =#007000>"
           ' Else
           '     If xYROPINF0.ROPINFIDT = mailYROPINF0_Suivant.ROPINFIDT And mailYROPINF0_Suivant.ROPINFIDT > 0 Then
           '         wTxt_BackColor = "<font color =#FF8000>"
           '     Else
           '         wTxt_BackColor = "<font color = #000000>"
           '     End If
           ' End If
           '___________________________________________________________________________________________
            ' If xYROPINF0.ROPINFIDT = mailYROPINF0.ROPINFIDT Then
           '     If xYROPINF0.ROPINFIDT2 = mailYROPINF0.ROPINFIDT2 Then wTxt_BackColor = "<font color =#007000>"
           ' Else
                    If xYROPINF0.ROPINFUAMJ = DSys Then
                        wTxt_BackColor = "<font color =#FF5000>"
                    Else
                        wTxt_BackColor = "<font color = #0000FF>"
                    End If
                'End If
           ' End If
       End If
        Select Case xYROPINF0.ROPINFGNAT
            Case "P":
                libROPINFGNAT = "* " & xYROPINF0.ROPINFIDP
                wTD_BackColor = "bgcolor = #90FFFF"
                wSta_BackColor = wTD_BackColor
                X = cmdSendMail_Sta(xYROPINF0.ROPINFSTA, wSta_BackColor)
                xUsr = Trim(cmdSendMail_USR(xYROPINF0.ROPINFGUSR))
                wDétail = wDétail & "<NOBR><TABLE  width=800 border=1   cellpadding=5 ></B><TR>" _
                         & "<TD width=5 height=10 Font color = #000000>" & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=195 height=10><Font size=2>" & mFontColor_Blue & libROPINFGNAT & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=400 height=10><Font size=2>" & mFontColor_Blue & xUsr & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=100 height=10><Font size=2>" & cmdSendMail_Ech(xYROPINF0.ROPINFGECH, xYROPINF0.ROPINFSTA) & "</TD>" _
                         & "<TD " & wSta_BackColor & " width=100 height=10><Font size=2>" & X & "</TD>" _
                         & "</TR>" _
                         & "<TR><TD Font color = #000000 >" & "</TD>" _
                         & "<TD colspan=4  height=10><PRE><Font size=2>" & wTxt_BackColor & cmdSendMail_Txt(xYROPINF0.ROPINFGTXT, "H") _
                         & "</TD></TR></TABLE>"
             Case "A", "F":
                wTD_BackColor = "bgcolor =  #E0FFFF"
                wSta_BackColor = wTD_BackColor
                X = cmdSendMail_Sta(xYROPINF0.ROPINFSTA, wSta_BackColor)
                If xYROPINF0.ROPINFIDT = mailYROPINF0_Suivant.ROPINFIDT And mailYROPINF0_Suivant.ROPINFIDT > 0 Then
                    wSta_BackColor = "bgcolor =#FFB000"
                End If

                xUsr = cmdSendMail_USR(xYROPINF0.ROPINFGUSR)
                wDétail = wDétail & "<NOBR><TABLE  width=800 border=1    cellpadding=5 ><TR>" _
                         & "<TD Font color = #000000 width=30  height=10><Font size=2>" & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=170  height=10><Font size=2>" & mFontColor_Blue & libROPINFGNAT & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=400  height=10><Font size=2>" & mFontColor_Blue & xUsr & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=100  height=10><Font size=2>" & cmdSendMail_Ech(xYROPINF0.ROPINFGECH, xYROPINF0.ROPINFSTA) & "</TD>" _
                         & "<TD " & wSta_BackColor & " width=100  height=10><Font size=2>" & X & "</TD>" _
                         & "</TR>" _
                         & "<TR><TD Font color = #000000 >" & "</TD>" _
                         & "<TD colspan=4  height=10><PRE><Font size=2>" & wTxt_BackColor & cmdSendMail_Txt(xYROPINF0.ROPINFGTXT, "H") _
                         & "</TD></TR></TABLE>"
              Case "J"
                    X = fraDétail_Display_PJ_FileName(xYROPINF0.ROPINFGUSR, False)
                    wSendMail.Attachment = wSendMail.Attachment & X & ";"
                    wTD_BackColor = "bgcolor = #F5F5F5"
                    wSta_BackColor = wTD_BackColor
                    X = cmdSendMail_Sta(xYROPINF0.ROPINFSTA, wSta_BackColor)
                    xUsr = cmdSendMail_USR(xYROPINF0.ROPINFGUSR)
                    wDétail = wDétail & "<TABLE   width=800 border=1 cellpadding=5 ><TR>" _
                         & "<TD Font color = #000000 width=30 height=10><Font size=2>" & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=170 height=10><Font size=2>" & mFontColor_Green & "PJ" & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=400 height=10><Font size=2>" & mFontColor_Green & Trim(xYROPINF0.ROPINFGTXT) & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=100 height=10><Font size=2>" & mFontColor_Green & dateImp10(xYROPINF0.ROPINFGECH) & "</TD>" _
                         & "<TD " & wSta_BackColor & " width=100 height=10><Font size=2>" & X & "</TD>" _
                         & "</TD></TR></TABLE>"
               Case Else
                    Select Case xYROPINF0.ROPINFGNAT
                        Case "N": libROPINFGNAT = "Note"
                        Case Else: libROPINFGNAT = xYROPINF0.ROPINFGNAT
                    End Select
                    wTD_BackColor = "bgcolor = #F5F5F5"
                    wSta_BackColor = wTD_BackColor
                    X = cmdSendMail_Sta(xYROPINF0.ROPINFSTA, wSta_BackColor)
                    xUsr = cmdSendMail_USR(xYROPINF0.ROPINFGUSR)
                    wDétail = wDétail & "<TABLE   width=800 border=1 cellpadding=5 ><TR>" _
                         & "<TD Font color = #000000 width=30 height=10><Font size=2>" & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=170 height=10><Font size=2>" & mFontColor_Green & libROPINFGNAT & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=400 height=10><Font size=2>" & mFontColor_Green & xUsr & "</TD>" _
                         & "<TD " & wTD_BackColor & " width=100 height=10><Font size=2>" & mFontColor_Green & dateImp10(xYROPINF0.ROPINFGECH) & "</TD>" _
                         & "<TD " & wSta_BackColor & " width=100 height=10><Font size=2>" & X & "</TD>" _
                         & "</TR>" _
                         & "<TR><TD Font color = #000000 height=10>" & "</TD>" _
                         & "<TD colspan=3  height=10><PRE><Font size=2>" & mFontColor_Green & cmdSendMail_Txt(xYROPINF0.ROPINFGTXT, "H") _
                         & "</TD></TR></TABLE>"
               End Select
            
    End If
Next K

'wSendMail.Message = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">"
'                    & "<FONT face=" & Asc34 & prtFontName_Arial & Asc34 & ">" _
'                    & wDétail _
'                    & "<BR>"
wSendMail.Message = wDétail
wSendMail.AsHTML = True
MsgBox "cmdsensmail Exit"
Exit Sub
If Not blnOff_Line And mailYROPINF0.ROPINFID > 2000 Then
    srvSendMail.Monitor wSendMail
    For K1 = 1 To arrRecipient_Nb
        lstErr.AddItem "@ " & arrRecipient(K1)
    Next K1

End If

End Sub


Public Sub fraSelect_Display()
Dim V
Dim X As String, X1 As String
fraSelect_Update.Visible = True
fraSelect_Update_B.Enabled = True
fraSelect_Display_Reset

fraDétail_Update_B.Enabled = False
chkSelect_Update_B.Value = "0"
Call lstErr_Clear(lstErr, cmdContext, ">Affichage du dossier"): DoEvents
libUpdate_ROPDOSID = Trim(xYROPDOS0.ROPDOSID) & " -" & xYROPDOS0.ROPDOSXID
libUpdate_ROPDOSIUSR = "=? " & dateImp10(xYROPDOS0.ROPDOSIAMJ) & " " & tvwSelect_Display_USR(xYROPDOS0.ROPDOSIUSR)
libUpdate_ROPDOSGUSR = "=> " & dateImp10(xYROPDOS0.ROPDOSGECH) & " " & tvwSelect_Display_USR(xYROPDOS0.ROPDOSGUSR)
libUpdate_ROPDOSUUSR = xYROPDOS0.ROPDOSUUSR _
                   & " " & dateImpS(xYROPDOS0.ROPDOSUAMJ) & " " & timeImp8(xYROPDOS0.ROPDOSUHMS) _
                   & "  v_" & xYROPINF0.ROPINFUVER
cbo_Scan xYROPDOS0.ROPDOSSTA, txtUpdate_ROPDOSSTA


cbo_Scan Trim(xYROPDOS0.ROPDOSXDOM), txtUpdate_ROPDOSXDOM
sqlYBIATAB0_cboID_K1 "ROPDOSXAPP", xYROPDOS0.ROPDOSXDOM, txtUpdate_ROPDOSXAPP
cbo_Scan Trim(xYROPDOS0.ROPDOSXAPP), txtUpdate_ROPDOSXAPP
txtUpdate_ROPDOSIREF = Trim(xYROPDOS0.ROPDOSIREF)
txtUpdate_ROPDOSXID = Trim(xYROPDOS0.ROPDOSXID)
txtUpdate_ROPDOSGCOU = Trim(xYROPDOS0.ROPDOSGCOU)

cbo_Scan Trim(xYROPDOS0.ROPDOSGUSR), txtUpdate_ROPDOSGUSR
Call DTPicker_Set(txtUpdate_ROPDOSGECH, xYROPDOS0.ROPDOSGECH)
cbo_Scan Trim(xYROPDOS0.ROPDOSIUSR), txtUpdate_ROPDOSIUSR
Call DTPicker_Set(txtUpdate_ROPDOSIAMJ, xYROPDOS0.ROPDOSIAMJ)
cbo_Scan xYROPDOS0.ROPDOSGNAT, txtUpdate_ROPDOSGNAT
cbo_Scan xYROPDOS0.ROPDOSGPRI, txtUpdate_ROPDOSGPRI
cbo_Scan xYROPDOS0.ROPDOSGPRV, txtUpdate_ROPDOSGPRV
cbo_Scan xYROPDOS0.ROPDOSGGRA, txtUpdate_ROPDOSGGRA
'If blnROPDOSQUAL Then
cbo_Scan xYROPDOS0.ROPDOSQUAL, txtUpdate_ROPDOSQUAL

End Sub
Public Function fraSelect_Update_Control()
Dim V
Dim K As Long

Dim blnUpdate_Control As Boolean
Dim X As String
blnUpdate_Control = True
Call lstErr_AddItem(lstErr, cmdContext, ">Contrôle dossier "): DoEvents
newYROPDOS0 = oldYROPDOS0
fraSelect_Display_Reset

Call DTPicker_Control(txtUpdate_ROPDOSGECH, newYROPDOS0.ROPDOSGECH)
If newYROPDOS0.ROPDOSGECH < DSys And newYROPDOS0.ROPDOSID > 2000 Then
    lblUpdate_ROPDOSGECH.BackColor = vbRed
    txtUpdate_ROPDOSGECH.ToolTipText = "L'échéance du dossier ne peut pas être < à aujourd'hui " & " | " & dateImp10(arrYROPINF0(K).ROPINFGECH)
    blnUpdate_Control = False
    Call lstErr_AddItem(lstErr, cmdContext, "?_________échéance < " & DSys)
Else
    For K = 2 To arrYROPINF0_Nb
        If newYROPDOS0.ROPDOSGECH < arrYROPINF0(K).ROPINFGECH Then
            lblUpdate_ROPDOSGECH.BackColor = vbRed
            txtUpdate_ROPDOSGECH.ToolTipText = "L'échéance du dossier ne peut pas être < à l'échéance de l'action " & arrYROPINF0(K).ROPINFIDP & "-" & arrYROPINF0(K).ROPINFIDT & " | " & dateImp10(arrYROPINF0(K).ROPINFGECH)
            blnUpdate_Control = False
            Call lstErr_AddItem(lstErr, cmdContext, "?_________éch Dossier < éch Action " & arrYROPINF0(K).ROPINFIDP & " / " & arrYROPINF0(K).ROPINFIDT)
            Exit For
        End If
    Next K
End If

Call DTPicker_Control(txtUpdate_ROPDOSIAMJ, newYROPDOS0.ROPDOSIAMJ)
If newYROPDOS0.ROPDOSIAMJ > newYROPDOS0.ROPDOSUAMJ Then
    lblUpdate_ROPDOSIAMJ.BackColor = vbRed
    txtUpdate_ROPDOSIAMJ.ToolTipText = "La date du constat ne peut pas être > à " & dateImp10(newYROPDOS0.ROPDOSUAMJ)
    blnUpdate_Control = False
    Call lstErr_AddItem(lstErr, cmdContext, "?_________date constat > date saisie")
End If

newYROPDOS0.ROPDOSGUSR = txtUpdate_ROPDOSGUSR
If Trim(newYROPDOS0.ROPDOSGUSR) = "?" Then
    If Not blnDossierModèle Then
        txtUpdate_ROPDOSGUSR.BackColor = vbRed
        blnUpdate_Control = False
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Gestionnaire ? ")
    End If
End If

newYROPDOS0.ROPDOSIUSR = txtUpdate_ROPDOSIUSR
If Trim(newYROPDOS0.ROPDOSIUSR) = "?" Then
    If Not blnDossierModèle Then
        txtUpdate_ROPDOSIUSR.BackColor = vbRed
        blnUpdate_Control = False
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Initiateur du constat ? ")
    End If
End If

newYROPDOS0.ROPDOSGNAT = txtUpdate_ROPDOSGNAT
If Trim(newYROPDOS0.ROPDOSGNAT) = "?" Then
    If Not blnDossierModèle Then
        txtUpdate_ROPDOSGNAT.BackColor = vbRed
        blnUpdate_Control = False
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Nature ? ")
    End If
End If

newYROPDOS0.ROPDOSGPRV = txtUpdate_ROPDOSGPRV
If Trim(newYROPDOS0.ROPDOSGPRV) = "?" Then
    If Not blnDossierModèle Then
        txtUpdate_ROPDOSGPRV.BackColor = vbRed
        blnUpdate_Control = False
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Confidentialité ? ")
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
    Else
        If Trim(newYROPDOS0.ROPDOSXDOM) = "Modèle" Then
            txtUpdate_ROPDOSXDOM.BackColor = vbRed
            txtUpdate_ROPDOSXDOM.ToolTipText = "Le domaine 'Modèle' n'est pas autorisé pour un dossier"
            blnUpdate_Control = False
            Call lstErr_AddItem(lstErr, cmdContext, "?_________'Modèle' n'est pas un domaine autorisé ? ")
        End If
    End If
End If
newYROPDOS0.ROPDOSXAPP = Mid$(txtUpdate_ROPDOSXAPP, 1, 12)
newYROPDOS0.ROPDOSIREF = txtUpdate_ROPDOSIREF
newYROPDOS0.ROPDOSXID = txtUpdate_ROPDOSXID
newYROPDOS0.ROPDOSGCOU = txtUpdate_ROPDOSGCOU
'If blnROPDOSQUAL Then
newYROPDOS0.ROPDOSQUAL = Mid$(txtUpdate_ROPDOSQUAL, 1, 3)

If newYROPDOS0.ROPDOSGPRV = "U" Then
    If Trim(newYROPDOS0.ROPDOSGUSR) <> usrName_UCase Then
        txtUpdate_ROPDOSGUSR.BackColor = vbRed
        txtUpdate_ROPDOSGUSR.ToolTipText = "Dossier privé(U) => le superviseur doit être = " & usrName_UCase
        blnUpdate_Control = False
        Call lstErr_AddItem(lstErr, cmdContext, "?_________dossier U => cet utilisateur n'est pas autorisé ")
    End If
End If


If blnUpdate_Control Then
    fraSelect_Update_Control = Null
Else
    fraSelect_Update_Control = "<Fin du contrôle des données "
End If
End Function

Public Function fraDétail_Update_Control()
Dim V
Dim blnUpdate_Control As Boolean
Dim X As String, lngX As Long
blnUpdate_Control = True
Call lstErr_AddItem(lstErr, cmdContext, ">Contrôle Information "): DoEvents
fraDétail_Display_Reset

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
        End If
        If newYROPINF0.ROPINFGECH > oldYROPDOS0.ROPDOSGECH Then
            lblUpdate_ROPINFGECH.BackColor = vbRed
            txtUpdate_ROPINFGECH.ToolTipText = "l'échéance ne peut être > à l'échéance du dossier " & dateImp10(oldYROPDOS0.ROPDOSGECH)
           blnUpdate_Control = False
            Call lstErr_AddItem(lstErr, cmdContext, "?_________échéance Info > éch Dossier")
        End If
    End If
End If

lngX = CLng(Val(txtUpdate_ROPINFGUO) * 100)
If lngX Mod 100 > 60 Then
    txtUpdate_ROPINFGUO.BackColor = vbRed
    txtUpdate_ROPINFGUO.ToolTipText = "la partie décimale (minute) doit être <= à 60"
    blnUpdate_Control = False
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Durée.Min > 60")
End If
If lngX > 99999 Then
    txtUpdate_ROPINFGUO.BackColor = vbRed
    txtUpdate_ROPINFGUO.ToolTipText = "le nombre d'heures ne peut pas être > à 999 h"
    blnUpdate_Control = False
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Durée.heure > 999")
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
End If
End Function

Public Sub fraDétail_Display()
Dim V, I As Integer, K As Integer
Dim X As String, X1 As String

blnControl = False
fraDétail_Display_Reset


libUpdate_ROPINFID = "dossier : " & xYROPINF0.ROPINFID & "   processus : " & xYROPINF0.ROPINFIDP _
                   & "   action : " & xYROPINF0.ROPINFIDT & "-" & xYROPINF0.ROPINFIDT2 _
                   & "   version : " & xYROPINF0.ROPINFUVER & "   mise à jour par : " & xYROPINF0.ROPINFUUSR _
                   & " le " & dateImp10(xYROPINF0.ROPINFUAMJ) & " " & timeImp8(xYROPINF0.ROPINFUHMS)
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
chkUpdate_ROPINFMAIL_D.Caption = tvwSelect_Display_USR(oldYROPDOS0.ROPDOSGUSR)

If Mid$(xYROPINF0.ROPINFMAIL, 1, 1) = "I" Then
    chkUpdate_ROPINFMAIL_I = "1"
Else
    chkUpdate_ROPINFMAIL_I = "0"
End If
chkUpdate_ROPINFMAIL_I.Caption = tvwSelect_Display_USR(oldYROPDOS0.ROPDOSIUSR)

If Mid$(xYROPINF0.ROPINFMAIL, 3, 1) = "P" Then
    chkUpdate_ROPINFMAIL_P = "1"
Else
    chkUpdate_ROPINFMAIL_P = "0"
End If
If Processus_Index > 0 Then chkUpdate_ROPINFMAIL_P.Caption = tvwSelect_Display_USR(arrYROPINF0(Processus_Index).ROPINFGUSR)

If Mid$(xYROPINF0.ROPINFMAIL, 4, 1) = "A" Then
    chkUpdate_ROPINFMAIL_A = "1"
Else
    chkUpdate_ROPINFMAIL_A = "0"
End If
If Action_Index > 0 Then
    chkUpdate_ROPINFMAIL_A.Caption = tvwSelect_Display_USR(xYROPINF0.ROPINFGUSR)
Else
    chkUpdate_ROPINFMAIL_A.Caption = ""
End If

chkUpdate_ROPINFMAIL_U = "0"
If Action_Suivante_Index > 0 Then
    chkUpdate_ROPINFMAIL_U.Caption = tvwSelect_Display_USR(arrYROPINF0(Action_Suivante_Index).ROPINFGUSR)
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
    X = fraDétail_Display_PJ_FileName(xYROPINF0.ROPINFGUSR, True)

End If

cmdSelect_Update.Enabled = True
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
blnControl = True

End Sub

Private Sub mnuPrint0_Liste_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
'cmdPrint_Facture

Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdSelect_Update_Init_04()
blnControl = False
mailSubject = "Création d'un dossier"
'$20071122_JPL$ chkSelect_Update_B.Value = "1"
fraDétail_Update_B.Enabled = True
'txtUpdate_ROPDOSGUSR.ListIndex = 0
txtUpdate_ROPDOSGTXT = ""
'cbo_Scan "D", txtUpdate_ROPDOSGNAT

txtUpdate_ROPINFGUSR.Enabled = True
txtUpdate_ROPINFGNAT.Enabled = False
txtUpdate_ROPINFGECH.Enabled = True 'False
If oldYROPDOS0.ROPDOSID >= 2000 Then
    chkUpdate_ROPINFMAIL_I = "0"
    chkUpdate_ROPINFMAIL_D = "0"
    chkUpdate_ROPINFMAIL_P = "0"
    chkUpdate_ROPINFMAIL_A = "0"
    chkUpdate_ROPINFMAIL_U = "0"
End If
chkUpdate_ROPINFGPRV = "1"
txtUpdate_ROPINFGTXT.Locked = False

'cbo_Scan Trim(usrName_UCase), txtUpdate_ROPINFGUSR
txtUpdate_ROPINFGNAT.Enabled = False
'Call DTPicker_Set(txtUpdate_ROPINFGECH, DSys)
txtUpdate_ROPINFGUO = "": txtUpdate_ROPINFGUO.Enabled = False
Call fraDétail_lbl("P", 1)
'libUpdate_ROPINFGTXT.Visible = True
chkSelect_Update_B.Value = "1"
blnSendMail = True
blnControl = True
End Sub

Private Sub cmdSelect_Update_Init_11()
mailSubject = "Modification d'une information"
fraDétail_Update_B.Enabled = True
txtUpdate_ROPINFGUSR.Enabled = True
txtUpdate_ROPINFGNAT.Enabled = False
txtUpdate_ROPINFGECH.Enabled = True
txtUpdate_ROPINFGUO.Enabled = True

If blnYROPINF0_12X Then
    txtUpdate_ROPINFGTXT.Locked = True
    libUpdate_ROPINFGTXT.Visible = False
Else
    txtUpdate_ROPINFGTXT.Locked = False
    libUpdate_ROPINFGTXT.Visible = True
End If

cmdSelect_Update_Ok.Visible = True
On Error Resume Next
txtUpdate_ROPINFGTXT.SetFocus
txtUpdate_ROPINFGTXT.SelStart = Len(txtUpdate_ROPINFGTXT)

End Sub

Private Sub cmdSelect_Update_Init_01()
blnControl = False
mailSubject = "Création d'une note"
fraDétail_Update_B.Enabled = True
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
cmdSelect_Update_Ok.Visible = True
txtUpdate_ROPINFGUO = "": txtUpdate_ROPINFGUO.Enabled = False
Call fraDétail_lbl("N", 0)
blnControl = True
End Sub


Private Sub cmdSelect_Update_Init_05()
blnControl = False
mailSubject = "Ajout d'une pièce jointe"
fraDétail_Update_B.Enabled = True
txtUpdate_ROPINFGUSR.Enabled = False
txtUpdate_ROPINFGNAT.Enabled = False
txtUpdate_ROPINFGECH.Enabled = False
'txtUpdate_ROPINFMAIL.Enabled = True

'cbo_Scan Trim(usrName_UCase), txtUpdate_ROPINFGUSR
cbo_Scan "J", txtUpdate_ROPINFGNAT
'cbo_Scan " ", txtUpdate_ROPINFMAIL
Call DTPicker_Set(txtUpdate_ROPINFGECH, DSys)
txtUpdate_ROPINFGTXT = ""

txtUpdate_ROPINFGTXT.Locked = True 'False
fraDétail_Update_J.Visible = True
blntvwSelect_Click = True
cmdSelect_Update_Ok.Visible = True
txtUpdate_ROPINFGUO = "": txtUpdate_ROPINFGUO.Enabled = False
oldFileName = "": newFileName = ""
blnControl = True
End Sub

Private Sub cmdSelect_Update_Init_02(lROPINFGNAT As String)
blnControl = False
mailSubject = "Création d'une action"
fraDétail_Update_B.Enabled = True
txtUpdate_ROPINFGUSR.Enabled = True
txtUpdate_ROPINFGNAT.Enabled = False
txtUpdate_ROPINFGECH.Enabled = True
txtUpdate_ROPINFGUO.Enabled = True
chkUpdate_ROPINFGPRV = "1"
    
If oldYROPDOS0.ROPDOSGPRV = "U" Then
    cbo_Scan Trim(usrName_UCase), txtUpdate_ROPINFGUSR
Else
    txtUpdate_ROPINFGUSR.ListIndex = 0
End If
cbo_Scan lROPINFGNAT, txtUpdate_ROPINFGNAT
Call DTPicker_Set(txtUpdate_ROPINFGECH, arrYROPINF0(Processus_Index).ROPINFGECH)
txtUpdate_ROPINFGTXT = ""

txtUpdate_ROPINFGTXT.Locked = False
cmdSelect_Update_Ok.Visible = True
txtUpdate_ROPINFGUO = "": txtUpdate_ROPINFGUO.Enabled = True
Call fraDétail_lbl("A", 0)

libUpdate_ROPINFGTXT.Visible = True
blnControl = True
End Sub
Private Sub cmdSelect_Update_Init_03()
blnControl = False
mailSubject = "Création d'un processus"
fraDétail_Update_B.Enabled = True
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
cmdSelect_Update_Ok.Visible = True
txtUpdate_ROPINFGUO = "": txtUpdate_ROPINFGUO.Enabled = False
Call fraDétail_lbl("P", 0)
blnControl = True
End Sub

Private Sub cmdSelect_Update_Init_14()
mailSubject = "Modification d'un dossier"
txtUpdate_ROPDOSGTXT.Visible = False: cmdSelect_Update_G.Visible = False
fraSelect_Update_B.Enabled = True
cmdSelect_Update_Ok.Visible = True
blnSelect_Update_B_Display = True
chkSelect_Update_B.Value = "1"
End Sub

Private Sub cmdSelect_Update_Init_31()
Dim X As String
mailSubject = "Annulation d'une note"
X = "Voulez-vous réellement ANNULER cette information?"
mROPINFSTA_Value = "A"
blntvwSelect_NodeClick = False
X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, libUpdate_ROPINFID.Caption)
If X = vbYes Then
    cmdDétail_Update_Ok
Else
    cmdSelect_Update_Reset
End If
End Sub

Private Sub cmdSelect_Update_Init_21()
Dim X As String
mailSubject = "Fermeture d'une note"
X = "Voulez-vous réellement Clôturer cette information?"
mROPINFSTA_Value = "F"
blntvwSelect_NodeClick = False
X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraDétail_Update_B)
If X = vbYes Then
    cmdDétail_Update_Ok
Else
    cmdSelect_Update_Reset
End If

End Sub

Private Sub cmdSelect_Update_Init_51()
Dim X As String
mailSubject = "Restauration d'une note"

X = "Voulez-vous réellement RESTAURER cette information?"
mROPINFSTA_Value = " "
blntvwSelect_NodeClick = True
X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraDétail_Update_B)
If X = vbYes Then
    cmdDétail_Update_Ok
Else
    cmdSelect_Update_Reset
End If

End Sub

Private Sub cmdSelect_Update_Init_52()
Dim X As String
mailSubject = "Restauration d'une action"

X = "Voulez-vous réellement RESTAURER cette action?"
mROPINFSTA_Value = " ": mROPINFSTA_Set = " ":: mROPINFSTAK_Set = " "
If oldYROPINF0.ROPINFSTA = "F" Then
    mROPINFSTA_Where = "+"
Else
    mROPINFSTA_Where = "-"
End If


X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraDétail_Update_B)
If X = vbYes Then
    cmdDétail_Update_Ok
Else
    cmdSelect_Update_Reset
End If
End Sub

Private Sub cmdSelect_Update_Init_53()
Dim X As String
mailSubject = "Restauration d'un processus"

X = "Voulez-vous réellement RESTAURER ce processus?"
mROPINFSTA_Value = " ": mROPINFSTA_Set = " ": mROPINFSTAK_Set = " "
If oldYROPINF0.ROPINFSTA = "F" Then
    mROPINFSTA_Where = "*"
Else
    mROPINFSTA_Where = "%"
End If


X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraDétail_Update_B)
If X = vbYes Then
    cmdDétail_Update_Ok
Else
    cmdSelect_Update_Reset
End If
End Sub


Private Sub cmdSelect_Update_Init_22()
Dim X As String
mailSubject = "Fermeture d'une action"

X = "Voulez-vous réellement Clôturer cette Action?"
mROPINFSTA_Value = "F": mROPINFSTA_Set = "+": mROPINFSTA_Where = " ": mROPINFSTAK_Set = "V"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraDétail_Update_B)
If X = vbYes Then
    If Mid$(cmdSelect_Update, 6, 1) = "+" Then    ' fermeture de l'action et création d'une autre action
        cmdDétail_Update_Ok
        arrSelect_Update_Nb = 1
        arrSelect_Update(1) = "02 - Ajouter une action"
        cmdSelect_Update_Display (arrSelect_Update(1))
        'cmdSelect_Update.Clear
        'cmdSelect_Update.AddItem "- 02 - Ajouter une action"
        cmdSelect_Update.ListIndex = 0
    Else
        cmdDétail_Update_Ok
    End If
Else
    cmdSelect_Update_Reset
End If
End Sub

Private Sub cmdSelect_Update_Init_23()
Dim X As String
mailSubject = "Fermeture d'un processus"

X = "Voulez-vous réellement Clôturer ce processus?"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraDétail_Update_B)
If X = vbYes Then
    cmdSelect_Update_Init_23_Ok
Else
    cmdSelect_Update_Reset
End If

End Sub
Private Sub cmdSelect_Update_Init_24()
Dim X As String
mailSubject = "Fermeture d'un dossier"

X = "Voulez-vous réellement Clôturer ce dossier?"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraDétail_Update_B)
If X = vbYes Then
    newYROPDOS0 = oldYROPDOS0
    cmdSelect_Update_Init_24_Ok
Else
    cmdSelect_Update_Reset
End If
End Sub

Private Sub cmdSelect_Update_Init_54()
Dim X As String
mailSubject = "Restauration d'un dossier"

X = "Voulez-vous réellement RESTAURER ce dossier?"
mROPINFSTA_Value = " ": mROPINFSTA_Set = " ": mROPINFSTAK_Set = " "
If oldYROPDOS0.ROPDOSSTA = "F" Then
    mROPINFSTA_Where = "$"
Else
    mROPINFSTA_Where = "£"
End If

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraDétail_Update_B)
If X = vbYes Then
    Call arrYROPINF0_SQL(oldYROPDOS0.ROPDOSID)
    xYROPDOS0 = oldYROPDOS0
    tvwSelect_STAK
    
    newYROPDOS0 = oldYROPDOS0
    newYROPDOS0.ROPDOSSTA = mROPINFSTA_Value
    newYROPDOS0.ROPDOSSTAK = xYROPDOS0.ROPDOSSTAK
    cmdSelect_Update_Ok_Transaction "Update"
    tvwSelect_NodeClick tvwSelect.Nodes("D" & Format$(oldYROPDOS0.ROPDOSID, "000000000"))
Else
    cmdSelect_Update_Reset
End If
End Sub

Private Sub cmdSelect_Update_Init_64()
Dim X As String
Dim V, K As Long
Dim wNb As Long
mailSubject = "Report d'échéance d'un dossier"

X = InputBox("Indiquer le nombre de jours de report (0 pour abandonner)?")
If Not IsNumeric(X) Then
    MsgBox ("abandon : saisie non valide")
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
App_Debug = "cmdSelect_Update_Init_64"
'-------------------------------------------------------

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If IsNull(V) Then
'________________________________________________________________________________
    For K = 1 To arrYROPINF0_Nb
        If arrYROPINF0(K).ROPINFSTA = " " Then
           newYROPINF0 = arrYROPINF0(K)
           newYROPINF0.ROPINFGECH = DateAdd_AMJ("d", wNb, arrYROPINF0(K).ROPINFGECH)
           V = sqlYROPINF0_Update(newYROPINF0, arrYROPINF0(K))
           If Not IsNull(V) Then Exit For
        End If
    Next K
End If
'________________________________________________________________________________
If Not IsNull(V) Then
    V = cnSAB_Transaction("Rollback")
Else
    V = cnSAB_Transaction("Commit")
    cmdSelect_Update_Ok_Transaction ("Update")
End If
    
tvwSelect_NodeClick tvwSelect.Nodes("D" & Format$(oldYROPDOS0.ROPDOSID, "000000000"))

'------------------------------------------
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Sub
Private Sub cmdSelect_Update_Init_74()
mailSubject = "Duplication d'un modèle"
txtUpdate_ROPDOSGTXT.Visible = False: cmdSelect_Update_G.Visible = False
fraDétail_Update_B.Enabled = True
cmdSelect_Update_Ok.Visible = False
lstUpdate_Modèle.Visible = True
End Sub

Private Sub cmdSelect_Update_Init_34()
Dim X As String
mailSubject = "Annulation d'un dossier"

X = "Voulez-vous réellement ANNULER ce dossier?"
mROPINFSTA_Value = "A": mROPINFSTA_Set = "£": mROPINFSTA_Where = " ": mROPINFSTAK_Set = "A"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraDétail_Update_B)
If X = vbYes Then
    newYROPDOS0 = oldYROPDOS0
    newYROPDOS0.ROPDOSSTA = mROPINFSTA_Value
    newYROPDOS0.ROPDOSSTAK = "A"
    cmdSelect_Update_Ok_Transaction "Update"
    tvwSelect_NodeClick tvwSelect.Nodes("D" & Format$(oldYROPDOS0.ROPDOSID, "000000000"))
Else
    cmdSelect_Update_Reset
End If
End Sub

Private Sub cmdSelect_Update_Init_32()
Dim X As String
mailSubject = "Annulation d'une action"

X = "Voulez-vous réellement ANNULER cette Action?"
mROPINFSTA_Value = "A": mROPINFSTA_Set = "-": mROPINFSTA_Where = " ": mROPINFSTAK_Set = "A"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraDétail_Update_B)
If X = vbYes Then
    cmdDétail_Update_Ok
Else
    cmdSelect_Update_Reset
End If
End Sub

Private Sub cmdSelect_Update_Init_33()
Dim X As String
mailSubject = "Annulation d'un processus"

X = "Voulez-vous réellement ANNULER ce processus?"
mROPINFSTA_Value = "A": mROPINFSTA_Set = "%": mROPINFSTA_Where = " ": mROPINFSTAK_Set = "A"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraDétail_Update_B)
If X = vbYes Then
    cmdDétail_Update_Ok
Else
    cmdSelect_Update_Reset
End If
End Sub

Private Sub cmdSelect_Update_Init_42()
Dim X As String
mailSubject = "Effacement d'une action"

X = "Voulez-vous EFFACER défitivement cette Action?"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraDétail_Update_B)
If X = vbYes Then
    mROPINFSTA_Set = "E"
    cmdSelect_Update_Fct = "Delete"
    cmdDétail_Update_Ok
Else
    cmdSelect_Update_Reset
End If
End Sub

Private Sub cmdSelect_Update_Init_43()
Dim X As String
mailSubject = "Effacement d'un processus"

X = "Voulez-vous EFFACER défitivement ce Processus?"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraDétail_Update_B)
If X = vbYes Then
    mROPINFSTA_Set = "E"
    cmdSelect_Update_Fct = "Delete"
    cmdDétail_Update_Ok
Else
    cmdSelect_Update_Reset
End If
End Sub
Private Sub cmdSelect_Update_Init_44()
Dim X As String
mailSubject = "Effacement d'un dossier"

X = "Voulez-vous EFFACER défitivement ce dossier ?"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraDétail_Update_B)
If X = vbYes Then
    cmdSelect_Update_Ok_Transaction "Delete"
    tvwSelect.Nodes.Remove ("D" & Format$(oldYROPDOS0.ROPDOSID, "000000000"))
Else
    cmdSelect_Update_Reset
End If
End Sub

Private Sub cmdSelect_Update_Init_41()
Dim X As String, K As Integer
mailSubject = "Effacement d'une note"
    
X = "Voulez-vous  EFFACER défitivement cette information?"

X = MsgBox(X & vbCrLf & vbCrLf & txtUpdate_ROPINFGTXT, vbYesNo + vbQuestion + vbDefaultButton2, fraDétail_Update_B)
If X = vbYes Then
    mROPINFSTA_Set = "E"
    cmdSelect_Update_Fct = "Delete"
    cmdDétail_Update_Ok
Else
    cmdSelect_Update_Reset
End If
End Sub





Private Sub mnuParam_Copy_Click()
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

Private Sub tvwSelect_Click()
If chkSelect_Update_B.Value <> "1" Then fraSelect_Update.Left = fraSelect_Update_Right
cmdPrint.Enabled = False
End Sub

Private Sub tvwSelect_Collapse(ByVal Node As MSComctlLib.Node)
Node.BackColor = vbWhite
End Sub

Private Sub tvwSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkSelect_Update_B = "1" Or libUpdate_ROPINFGTXT.Visible Or lstUpdate_ROPINFMAIL.Visible Or fraDétail_Update_J.Visible Then
Else
    If blntvwSelect_Click Then
        fraSelect_Update.Left = fraSelect_Update_Left
    Else
        If fraSelect_Update.Left < fraSelect_Update_Right Then fraSelect_Update.Left = fraSelect_Update_Right
    End If
End If

End Sub


Private Sub tvwSelect_NodeClick(ByVal Node As MSComctlLib.Node)
Dim lenX As Integer, K As Integer
Dim X As String, xSql As String
Dim xKey As String, xParent As String
Dim xNode As Node
Dim blnDossier_Click As Boolean
Me.Enabled = False

If Not mSelect_Node Is Nothing Then
    tvwSelect.Nodes(Mid$(mSelect_Node.key, 1, 10)).Expanded = False
End If
If Not blnmSelect_Node Then
    Set mSelect_Node = Node
    Set xNode_Parent = Node.Parent
    arrYROPINF0_Index = 0
End If

lenX = Len(Node.key)
blnDossier_Click = False
If lenX <= 11 Then
    blnDossier_Click = True
    oldYROPDOS0.ROPDOSID = Mid$(mSelect_Node.key, 2, 9)
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 where ROPDOSID =" & oldYROPDOS0.ROPDOSID
    Set rsSab = cnsab.Execute(xSql)
    
    If Not rsSab.EOF Then
        V = rsYROPDOS0_GetBuffer(rsSab, oldYROPDOS0)
        arrYROPINF0_SQL oldYROPDOS0.ROPDOSID
        If Not blnmSelect_Node Then arrYROPINF0_Index = 1
    End If
  '  X = "D" & Format$(oldYROPDOS0.ROPDOSID, "000000000") & Format$(arrYROPINF0(arrYROPINF0_Nb).ROPINFIDP, "00000") & Format$(arrYROPINF0(arrYROPINF0_Nb).ROPINFIDT, "00000") & "00001"
  '  Set xNode = tvwSelect.Nodes(X)
  '  tvwSelect_NodeClick xNode
Else
    For arrYROPINF0_Index = 1 To arrYROPINF0_Nb
        xYROPINF0 = arrYROPINF0(arrYROPINF0_Index)
        xKey = "D" & Format$(xYROPINF0.ROPINFID, "000000000") & Format$(xYROPINF0.ROPINFIDP, "00000") & Format$(xYROPINF0.ROPINFIDT, "00000") & Format$(xYROPINF0.ROPINFIDT2, "00000")
        If xKey = mSelect_Node.key Then Exit For
    Next arrYROPINF0_Index
End If

blnmSelect_Node = False

If cmdSelect_SQL_K = "2" Then
    tvwSelect_Display
    fraSelect_Update.Left = fraSelect_Update_Left
    fraDétail_Update_B.Visible = False
    txtUpdate_ROPDOSGTXT.Visible = True: cmdSelect_Update_G.Visible = True
    X = MsgBox("Voulez-vous créer un nouveau dossier ?", vbYesNo + vbQuestion, Trim(oldYROPDOS0.ROPDOSXDOM) & " - " & Trim(oldYROPDOS0.ROPDOSXAPP))
    If X = vbYes Then
        fraDétail_Update_B.Visible = True
        cmdSelect_SQL_2_Duplication
    Else
            fraSelect_Update.Visible = False
    End If
    
Else
    tvwSelect_Display
    blntvwSelect_Click = True
    If blnDossier_Click Then
        'fraSelect_Update.Left = fraSelect_Update_Right
        'txtUpdate_ROPDOSGTXT.Visible = True
        blnSelect_Update_B_Display = False
            

        '$20071122_JPL$ chkSelect_Update_B.Value = "1"
    'Else
    '    blntvwSelect_Click = True
    End If
    If lenX <= 11 Then
        chkSelect_Update_B.Value = "1"
        If lenX = 11 Then Node.BackColor = &HE0E0E0
'        txtUpdate_ROPDOSGTXT.Visible = True
    End If
End If

Me.Enabled = True
End Sub

Private Sub tvwSelect_Display()
Dim lenX As Integer, K As Integer
Dim X As String, xSql As String
Dim xKey As String, xParent As String, xDossier_Key As String
Dim blnDisplay As Boolean, xDisplay_Séparation As String, xDisplay_Marque As String, xDisplay_Tabulation As String
Dim xProcessus As String
Dim Xdisplay As String * 120

'___________________________________________________________________________________
On Error Resume Next
xYROPDOS0 = oldYROPDOS0

tvwSelect_STAK
xDossier_Key = "D" & Format$(oldYROPDOS0.ROPDOSID, "000000000")
tvwSelect.Nodes.Remove (xDossier_Key)
If oldYROPDOS0.ROPDOSID < 1000 Then
    Set xNode = tvwSelect.Nodes.Add(, , xDossier_Key, Format$(oldYROPDOS0.ROPDOSID, "0000") & " - " & tvwSelect_ROPDOSXAPP_Libellé(oldYROPDOS0.ROPDOSXDOM, oldYROPDOS0.ROPDOSXAPP))
Else
    Set xNode = tvwSelect.Nodes.Add(, , xDossier_Key, Format$(oldYROPDOS0.ROPDOSID, "0000") & " - " & oldYROPDOS0.ROPDOSXDOM & " " & oldYROPDOS0.ROPDOSXAPP & " (" & Trim(tvwSelect_Display_USR(xYROPDOS0.ROPDOSGUSR)) & " | " & Trim(tvwSelect_Display_USR(xYROPDOS0.ROPDOSIUSR)) & ")")
End If
xNode.Bold = True
xNode.Expanded = True
tvwSelect_YROPDOS0_Forecolor
xNode.BackColor = &HC0F0FF
fraSelect_Display
'___________________________________________________________________________________
        
For K = 1 To arrYROPINF0_Nb
    xYROPINF0 = arrYROPINF0(K)
    If xYROPINF0.ROPINFSTA <> "E" Then
        xKey = xDossier_Key & Format$(xYROPINF0.ROPINFIDP, "00000") & Format$(xYROPINF0.ROPINFIDT, "00000") & Format$(xYROPINF0.ROPINFIDT2, "00000")
        If xYROPINF0.ROPINFIDT2 = 1 Then
            If xYROPINF0.ROPINFIDT = 0 Then
                xParent = xDossier_Key
            Else
                xParent = xDossier_Key & Format$(xYROPINF0.ROPINFIDP, "00000") & "0000000001"
            End If
        Else
            xParent = xDossier_Key & Format$(xYROPINF0.ROPINFIDP, "00000") & Format$(xYROPINF0.ROPINFIDT, "00000") & "00001"
                      
        End If
        If xYROPINF0.ROPINFGNAT = "P" Then
            xProcessus = Format(xYROPINF0.ROPINFIDP, "00") & " : "
        Else
            xProcessus = ""
        End If
        Xdisplay = xYROPINF0.ROPINFGTXT
        Set xNode = tvwSelect.Nodes.Add(xParent, tvwChild, xKey, xProcessus & dateImpS(xYROPINF0.ROPINFGECH) & " : " & tvwSelect_Display_USR(xYROPINF0.ROPINFGUSR) & " : " & Xdisplay)
        tvwSelect_YROPINF0_Forecolor
        xNode.BackColor = &HC0F0FF
        xNode.Expanded = True
    End If
Next K
'___________________________________________________________________________________


tvwSelect.Nodes(mSelect_Node.key).BackColor = vbCyan
tvwSelect.Nodes(mSelect_Node.key).Selected = True
If Not xNode_Parent Is Nothing Then xNode_Parent.Expanded = True

oldYROPINF0 = arrYROPINF0(arrYROPINF0_Index)
xYROPINF0 = oldYROPINF0

If mSelect_Node.key = xDossier_Key Or mSelect_Node.key = xDossier_Key & "+" Then
    cmdSelect_Update_Init_K = "D"
Else
    Select Case oldYROPINF0.ROPINFGNAT
        Case "P": cmdSelect_Update_Init_K = "P"
        Case "A", "F": cmdSelect_Update_Init_K = "A"
        Case Else: cmdSelect_Update_Init_K = " "
    End Select
End If
'___________________________________________________________________________________
Processus_Index = 0: Action_Index = 0: Action_Suivante_Index = 0
arrYROPINF0(0) = zYROPINF0
mailYROPINF0_Suivant = zYROPINF0

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
    Else
'         If xYROPINF0.ROPINFIDP > arrYROPINF0(K).ROPINFIDP Then Exit For
    End If
Next K
'___________________________________________________________________________________

fraDétail_Display
'___________________________________________________________________________________

txtUpdate_ROPDOSGTXT.ForeColor = vbBlack '&H606060

txtUpdate_ROPDOSGTXT = "Domaine         : " & Trim(Mid$(txtUpdate_ROPDOSXDOM, 1, 12)) & " - " & Trim(Mid$(txtUpdate_ROPDOSXAPP, 1, 12)) & vbCrLf _
                    & "Réf BIA /CRI    : " & txtUpdate_ROPDOSIREF & vbTab & "- " & txtUpdate_ROPDOSXID & vbCrLf _
                    & "Nature          : " & txtUpdate_ROPDOSGNAT & vbCrLf _
                    & "Priorité        : " & txtUpdate_ROPDOSGPRI & vbCrLf _
                    & "Gravité         : " & txtUpdate_ROPDOSGGRA & "    ( " & Trim(txtUpdate_ROPDOSGCOU) & " €)" & vbCrLf _
                    & "Confidentialité : " & txtUpdate_ROPDOSGPRV & vbCrLf
For K = 1 To arrYROPINF0_Nb
    blnDisplay = True 'False
    If arrYROPINF0(K).ROPINFGNAT = "P" Then
        If arrYROPINF0(K).ROPINFIDP = 1 And arrYROPINF0(K).ROPINFIDT = 0 Then blnDisplay = True
        If arrYROPINF0(K).ROPINFIDP = xYROPINF0.ROPINFIDP Then blnDisplay = True
    Else
        If arrYROPINF0(K).ROPINFIDP = xYROPINF0.ROPINFIDP Then blnDisplay = True
    End If
    
    If blnDisplay Then
        Select Case arrYROPINF0(K).ROPINFGNAT
            Case "P":      xDisplay_Séparation = vbCrLf & String(89, "="): xDisplay_Marque = "."
            'Case "A", "F": xDisplay_Séparation = String(88, "_")
            Case Else:     xDisplay_Séparation = String(89, "_"): xDisplay_Marque = "."
        End Select
        If arrYROPINF0(K).ROPINFSTA = " " Then xDisplay_Marque = ">"
        
  '      If arrYROPINF0(K).ROPINFSTA <> " " Then xDisplay_Marque = "  ": xDisplay_Séparation = String(62, "_") '"~"
        
  '      txtUpdate_ROPDOSGTXT = txtUpdate_ROPDOSGTXT & xDisplay_Séparation & xDisplay_Marque _
  '                       & dateImp10(arrYROPINF0(K).ROPINFGECH) & "  " & Trim(tvwSelect_Display_USR(arrYROPINF0(K).ROPINFGUSR)) _
  '                       & vbCrLf & vbCrLf & cmdSendMail_Txt(arrYROPINF0(K).ROPINFGTXT, "T") & vbCrLf & vbCrLf
         txtUpdate_ROPDOSGTXT = txtUpdate_ROPDOSGTXT & xDisplay_Séparation & xDisplay_Marque _
                         & vbCrLf & cmdSendMail_Txt(arrYROPINF0(K).ROPINFGTXT, "T") & vbCrLf
   End If
Next K
cmdSelect_Update_Init
blnSelect_Update_EnCours = False
blnSelect_Update_Expandable = True
'blntvwSelect_Click = False
End Sub
Public Function tvwSelect_Display_USR(lK1 As String) As String
Dim K As Integer, X As String

If Mid$(lK1, 1, 1) <> "_" Then
    tvwSelect_Display_USR = Format(lK1, "@")
Else
    K = Val(Mid$(lK1, 3, 2))
    tvwSelect_Display_USR = Format(arrROPDOSISRV_Code(K), "@")
End If
End Function

Public Function tvwSelect_Display_USR_Name(lK1 As String) As String
Dim K As Integer, X As String

If Mid$(lK1, 1, 1) <> "_" Then
    tvwSelect_Display_USR_Name = lK1
Else
    K = Val(Mid$(lK1, 3, 2))
    tvwSelect_Display_USR_Name = arrROPDOSISRV_Lib(K)
End If
End Function

Private Sub tvwSelect_STAK()

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
                Case "F": arrYROPINF0(K).ROPINFSTAK = "V": blnAction_Valide = True  '"F", "+", "$", "*"
                Case "A": arrYROPINF0(K).ROPINFSTAK = "A"  '"A", "-", "%", "£"
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
            Call tvwSelect_STAK_Processus(kProcessus)
            kProcessus = K
 '____________________________________________________________________________Autres
        Case Else:
            arrYROPINF0(K).ROPINFSTAK = " "
    End Select
Next K
'_________________________________________________________________________Dossier

Call tvwSelect_STAK_Processus(kProcessus)

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

Private Sub txtSelect_ROPDOSGUSR_GotFocus()
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
        cmdSelect_Update_K = 0
        tvwSelect_STAK
    If oldYROPDOS0.ROPDOSSTAK <> xYROPDOS0.ROPDOSSTAK Then
        newYROPDOS0 = xYROPDOS0
        cmdSelect_Update_Ok_Transaction ("Update")
    End If
'_________________________________________________________________________
    cmdSelect_SQL_6_Transaction
'_________________________________________________________________________

Next arrYROPDOS0_Index

End Sub
 
Public Sub cmdSelect_SQL_JPL()
Dim V, xSql As String
Dim xWhere As String
Dim K As Long
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_6"): DoEvents

'20071029 - ajout ROPDOSGSRV et ROPINFGSRV
'_______________________________________________
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then Exit Sub

'________________________________________________________________________________

xSql = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0  order by ROPDOSID"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYROPDOS0_GetBuffer(rsSab, oldYROPDOS0)

     If IsNull(V) Then
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "Màj DOS : " & oldYROPDOS0.ROPDOSID): DoEvents
        newYROPDOS0 = oldYROPDOS0
        Call cmdSelect_Update_Ok_GSRV(newYROPDOS0.ROPDOSIUSR, newYROPDOS0.ROPDOSISRV)
        Call cmdSelect_Update_Ok_GSRV(newYROPDOS0.ROPDOSGUSR, newYROPDOS0.ROPDOSGSRV)
        V = sqlYROPDOS0_Update(newYROPDOS0, oldYROPDOS0)
    End If
    rsSab.MoveNext

Loop
V = cnSAB_Transaction("Commit")
'_________________________________________________________________________
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then Exit Sub

xSql = "select * from " & paramIBM_Library_SABSPE & ".YROPINF0" _
     & " order by ROPINFID,ROPINFIDP,ROPINFIDT,ROPINFIDT2"
     
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYROPINF0_GetBuffer(rsSab, oldYROPINF0)

     If IsNull(V) Then
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "Màj INF : " & oldYROPINF0.ROPINFID): DoEvents
        newYROPINF0 = oldYROPINF0
        If oldYROPINF0.ROPINFGNAT = "J" Or oldYROPINF0.ROPINFGNAT = "M" Then
        Else
            Call cmdSelect_Update_Ok_GSRV(newYROPINF0.ROPINFGUSR, newYROPINF0.ROPINFGSRV)
            V = sqlYROPINF0_Update(newYROPINF0, oldYROPINF0)
        End If

    End If
    rsSab.MoveNext

Loop
'_________________________________________________________________________
V = cnSAB_Transaction("Commit")

End Sub

Public Sub cmdSelect_SQL_7()
Dim V, Nb As Long, NbD As Long, wROPINFID As Long
Dim xWhere As String, xSql As String

'Call DTPicker_Control(txtSelect_ROPDOSGECH_Max, wAmjMax)
    
xWhere = " where ROPINFSTA = ' ' and ROPINFID > 1000 and  ROPINFIDT > 0 and  ROPINFIDT2 = 1 and ROPINFGECH <= '" & DSys_SuivantO & "'"

xSql = "select  count(*)  as Tally from " & paramIBM_Library_SABSPE & ".YROPINF0 " & xWhere
Set rsSab = cnsab.Execute(xSql)
Nb = rsSab("Tally")
ReDim arrYROPINF0(Nb + 10)

arrYROPINF0_Nb = 0: NbD = 0: wROPINFID = 0
xSql = "select *  from " & paramIBM_Library_SABSPE & ".YROPINF0 " & xWhere _
     & " order by ROPINFID , ROPINFIDP , ROPINFIDT"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    V = rsYROPINF0_GetBuffer(rsSab, xYROPINF0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmROPINF.cmdSelect_SQL_7"
     Else
         arrYROPINF0_Nb = arrYROPINF0_Nb + 1
         arrYROPINF0(arrYROPINF0_Nb) = xYROPINF0
         If wROPINFID <> xYROPINF0.ROPINFID Then
            wROPINFID = xYROPINF0.ROPINFID
            NbD = NbD + 1
        End If
    End If
    rsSab.MoveNext
Loop

ReDim arrYROPDOS0(NbD + 10)
lstW.Clear

arrYROPDOS0_Nb = 0: wROPINFID = 0
For arrYROPINF0_Index = 1 To arrYROPINF0_Nb
    xYROPINF0 = arrYROPINF0(arrYROPINF0_Index)
    If wROPINFID <> xYROPINF0.ROPINFID Then
        wROPINFID = xYROPINF0.ROPINFID
        xSql = "select *  from " & paramIBM_Library_SABSPE & ".YROPDOS0 where ROPDOSID = " & xYROPINF0.ROPINFID
        Set rsSab = cnsab.Execute(xSql)
        If rsSab.EOF Then
            rsYROPDOS0_Init xYROPDOS0
             MsgBox V, vbCritical, "frmROPdos.cmdSelect_SQL_7.2"
        Else
            V = rsYROPDOS0_GetBuffer(rsSab, xYROPDOS0)
        End If
         If Not IsNull(V) Then
             MsgBox V, vbCritical, "frmROPdos.cmdSelect_SQL_7.3"
         Else
             arrYROPDOS0_Nb = arrYROPDOS0_Nb + 1
             arrYROPDOS0(arrYROPDOS0_Nb) = xYROPDOS0
        End If
    End If
    X = "_" & Format$(xYROPINF0.ROPINFID, "000000000") _
      & "_" & Format$(xYROPINF0.ROPINFIDP, "000000000") _
      & "_" & Format$(xYROPINF0.ROPINFIDT, "000000000") _
      & "_" & Format$(arrYROPINF0_Index, "000000000") _
      & "_" & Format$(arrYROPDOS0_Nb, "000000000") _
      & "_" & xYROPINF0.ROPINFGECH _
      & "_" & xYROPINF0.ROPINFGUSR
    
    cmdSelect_SQL_7_Destinataire
            
Next arrYROPINF0_Index
lstW.Visible = True

cmdPrint0_Echéancier

Exit Sub

End Sub

Private Sub txtSelect_ROPDOSGUSR_LostFocus()
txt_LostFocus txtSelect_ROPDOSGUSR

End Sub

Private Sub txtSelect_ROPDOSID_GotFocus()
txt_GotFocus txtSelect_ROPDOSID
End Sub


Private Sub txtSelect_ROPDOSID_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub

Private Sub txtSelect_ROPDOSID_LostFocus()
txt_LostFocus txtSelect_ROPDOSID

End Sub


Private Sub txtSelect_ROPDOSSTA_GotFocus()
txt_GotFocus txtSelect_ROPDOSSTA

End Sub


Private Sub txtSelect_ROPDOSSTA_LostFocus()
txt_LostFocus txtSelect_ROPDOSSTA

End Sub


Private Sub txtSelect_ROPDOSXAPP_GotFocus()
txt_GotFocus txtSelect_ROPDOSXAPP

End Sub


Private Sub txtSelect_ROPDOSXAPP_LostFocus()
txt_LostFocus txtSelect_ROPDOSXAPP

End Sub


Private Sub txtSelect_ROPDOSXDOM_Click()
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
txt_GotFocus txtSelect_ROPDOSXID

End Sub


Private Sub txtSelect_ROPDOSXID_LostFocus()
txt_LostFocus txtSelect_ROPDOSXID
End Sub


Private Sub txtSelect_ROPINFGTXT_GotFocus()
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
lblUpdate_ROPDOSGECH.BackColor = &HC0F0FF

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
fraDétail_Update_B.Enabled = False

txtUpdate_ROPINFGTXT.Locked = True
'cmdDétail_Update_Ok.Visible = False

Call fraDétail_lbl(oldYROPINF0.ROPINFGNAT, oldYROPINF0.ROPINFIDP)

blnDétail_Update_INF = False


If oldYROPDOS0.ROPDOSUUSR = usrName_UCase Or oldYROPINF0.ROPINFUUSR = usrName_UCase Then
    blnDétail_Update_INF = ROPDOS_Aut.Saisir
Else
    If ROPDOS_Aut.Xspécial Then
        blnDétail_Update_INF = True
    End If
End If


End Sub


Public Sub tvwSelect_YROPINF0_Forecolor()

xNode.BackColor = vbWhite
Select Case xYROPINF0.ROPINFSTA
    Case " ":    xNode.ForeColor = vbBlack ' '&H800000
    Case "A":    xNode.ForeColor = &H808080
    Case Else: xNode.ForeColor = &H4000&
End Select

Select Case xYROPINF0.ROPINFGNAT
    Case "A"
        Select Case xYROPINF0.ROPINFSTAK
            Case "V": xNode.IMAGE = "A_Vert"
            Case "R": xNode.IMAGE = "A_Rouge": xNode.ForeColor = vbRed
            Case "O": xNode.IMAGE = "A_Orange"
            Case "B": xNode.IMAGE = "A_Bleu"
            Case "A": xNode.IMAGE = "Non"
            Case "!": xNode.IMAGE = "Attention": xNode.ForeColor = vbMagenta
            Case Else: xNode.IMAGE = "A"
        End Select
    Case "P"
        'xNode.ForeColor = vbWhite
        Select Case xYROPINF0.ROPINFSTAK
            Case "V": xNode.IMAGE = "P_Vert": xNode.ForeColor = &H4000&
            Case "R": xNode.IMAGE = "P_Rouge": xNode.ForeColor = &H8080FF
            Case "O": xNode.IMAGE = "P_Orange": xNode.ForeColor = vbBlack ' '&H80FF&
            Case "B": xNode.IMAGE = "P_Bleu": xNode.ForeColor = vbBlack ' '&HC0C000
            Case "M": xNode.IMAGE = "P_Magenta": xNode.ForeColor = &HFF00FF
            Case "A": xNode.IMAGE = "Non": xNode.ForeColor = &HC0C0C0
            Case "!": xNode.IMAGE = "Attention": xNode.ForeColor = &HFF00FF
            Case Else: xNode.IMAGE = "P": xNode.ForeColor = vbBlack ' '&H808080
        End Select
    Case "F"
        Select Case xYROPINF0.ROPINFSTAK
            Case "V": xNode.IMAGE = "F_Vert"
            Case "R": xNode.IMAGE = "F_Rouge": xNode.ForeColor = vbRed
            Case "O": xNode.IMAGE = "F_Orange"
            Case "B": xNode.IMAGE = "F_Bleu"
            Case "A": xNode.IMAGE = "Non"
            Case "!": xNode.IMAGE = "Attention": xNode.ForeColor = vbMagenta
            Case Else: xNode.IMAGE = "F"
        End Select
    Case "J": xNode.IMAGE = "Trombon1": xNode.ForeColor = vbBlack ' '&H606060
    Case "N": xNode.IMAGE = "Note": xNode.ForeColor = vbBlack ' '&H606060
    Case Else
End Select
End Sub

Public Sub tvwSelect_YROPDOS0_Forecolor()

Select Case xYROPDOS0.ROPDOSSTA
    Case " ":    xNode.ForeColor = vbBlue '&HF00000
    Case "A":    xNode.ForeColor = &H808080
    Case Else: xNode.ForeColor = &H8000&
End Select

Select Case xYROPDOS0.ROPDOSSTAK
    Case "V": xNode.IMAGE = "Vert"
    Case "B": xNode.IMAGE = "Bleu"
    Case "O": xNode.IMAGE = "Orange"
    Case "R": xNode.IMAGE = "Rouge"
    Case "A": xNode.IMAGE = "Non"
    Case "!": xNode.IMAGE = "Attention"
    Case "M": xNode.IMAGE = "GrandStroumpf"
    Case Else: xNode.IMAGE = "D"
End Select
End Sub

Public Sub cmdSelect_Update_Init()
Dim blnROPINFGPRV As Boolean
Dim wUsr_Aut As String
Dim arrSelect_Update_K As Integer
Dim K As Integer, I As Integer
Dim blnDossierSAnsModèle As Boolean
Dim blnProcessusAvecActionFinale As Boolean

arrSelect_Update_K = 1

cmdSelect_Update.Clear: cmdSelect_Update_G.Clear
cmdSelect_Update.BackColor = &HC0FFC0    '&HFFC0FF
cmdSelect_Update.ForeColor = &H4000&     'vbBlue
libUpdate_ROPINFGTXT.Visible = False
lstUpdate_ROPINFMAIL.Visible = False
blnYROPINF0_12X_Aut = False
blnROPINFIDT_Insérer = False
blntvwSelect_NodeClick = True
cmdSelect_Update_Fct = "Update"
cmdSelect_Update_Ok.Caption = "Enregistrer"
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
If ROPDOS_Aut.Xspécial Then
    Dossier_Aut.Saisir = True
    Dossier_Aut.Valider = True
    Processus_Aut = Dossier_Aut
    Action_Aut = Dossier_Aut
    Memo_Aut = Dossier_Aut
Else

    If cmdSelect_Update_Init_00_Hab(Trim(oldYROPDOS0.ROPDOSGUSR)) Then
        Dossier_Aut.Saisir = True
        Dossier_Aut.Valider = True
        Processus_Aut = Dossier_Aut
        Action_Aut = Dossier_Aut
        Memo_Aut = Dossier_Aut
    Else
        If cmdSelect_Update_Init_00_Hab(Trim(arrYROPINF0(Processus_Index).ROPINFGUSR)) Then
            Processus_Aut.Saisir = True
            Processus_Aut.Valider = True
            Action_Aut = Processus_Aut
            Memo_Aut = Processus_Aut
        Else
            If cmdSelect_Update_Init_00_Hab(Trim(arrYROPINF0(Action_Index).ROPINFGUSR)) Then
                Action_Aut.Saisir = True
                Action_Aut.Valider = True
                Memo_Aut = Action_Aut
            Else
                If cmdSelect_Update_Init_00_Hab(Trim(oldYROPINF0.ROPINFGUSR)) Then
                    Memo_Aut.Saisir = True
                    Memo_Aut.Valider = True
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
    If arrYROPINF0(Processus_Index).ROPINFSTA <> " " Then
        Action_Aut = False_Aut
        Memo_Aut = False_Aut
    Else
        If arrYROPINF0(Action_Index).ROPINFSTA <> " " Then
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
arrSelect_Update_Nb = 1: arrSelect_Update(arrSelect_Update_Nb) = "00 --------------------------------------"
arrSelect_Update_Nb = 2: arrSelect_Update(arrSelect_Update_Nb) = "99 - Imprimer"
arrSelect_Update_Nb = 3: arrSelect_Update(arrSelect_Update_Nb) = "98 - envoyer le dossier par mail"
Select Case cmdSelect_Update_Init_K
    Case "D"
        cmdSelect_Update_G.AddItem "* 98 - envoyer le dossier par mail"
        cmdSelect_Update_G.AddItem "* 99 - Imprimer"
              If oldYROPDOS0.ROPDOSSTA = " " Then
              
                If Dossier_Aut.Saisir Then
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "03 - Ajouter un processus"
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "01 - Ajouter une note"
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "02 - Ajouter une action"
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "05 - Ajouter une pièce jointe"
                    cmdSelect_Update_G.AddItem "* 03 - Ajouter un processus"
              End If
               If Dossier_Aut.Valider Then
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "14 - modifier ce dossier "
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "24 - Clôturer ce dossier"
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "34 - Annuler ce dossier"
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "64 - Report d'échéance du dossier"
                    If blnDossierSAnsModèle Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "74 - choisir un modèle de gestion"
                    cmdSelect_Update_G.AddItem "* 14 - modifier ce dossier"
                    cmdSelect_Update_G.AddItem "* 24 - Clôturer ce dossier"
                    cmdSelect_Update_G.AddItem "* 64 - Report d'échéance du dossier"
            End If
               Else
                If Dossier_Aut.Valider Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "54 - Réactiver ce dossier"
            End If
            If Dossier_Aut.Xspécial Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "44 - Effacer ce dossier"
    Case "P"
             If xYROPINF0.ROPINFSTA = " " Then
                If Processus_Aut.Saisir Then
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "01 - Ajouter une note"
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "02 - Ajouter une action"
                    If Not blnProcessusAvecActionFinale Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "02F- Ajouter une action pour la fermeture du processus"
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "05 - Ajouter une pièce jointe"
                    If blnROPINFGPRV Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "13 - Modifier ce processus"
                End If
                If Processus_Aut.Valider Then
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "23 - Clôturer ce processus"
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "33 - Annuler ce processus"
                End If
             Else
                If Processus_Aut.Valider Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "53 - Réactiver ce processus"
            End If
           If Processus_Aut.Xspécial Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "43 - Effacer ce processus"
  Case "A"
            If xYROPINF0.ROPINFSTA = " " Then
                If Action_Aut.Saisir Then
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "01 - Ajouter une note"
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "02 - Ajouter une action"
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "02I- Insérer une action avant cette action"
                    arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "05 - Ajouter une pièce jointe"
                    If blnROPINFGPRV Then
                        arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "12 - Modifier cette action": arrSelect_Update_K = arrSelect_Update_Nb
                        blnYROPINF0_12X_Aut = True
                        arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "12X- Réaffecter(Resp,Ech) cette action"
                    End If
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
    Case Else
            If Memo_Aut.Saisir Then
                arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "01 - Ajouter une note"
                arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "05 - Ajouter une pièce jointe"
                If blnROPINFGPRV Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "11 - Modifier cette note": arrSelect_Update_K = arrSelect_Update_Nb
           End If
            If Memo_Aut.Valider And xYROPINF0.ROPINFSTA = " " Then
                arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "21 - Clôturer cette note"
                arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "31 - Annuler cette note"
            Else
                If Memo_Aut.Valider Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "51 - Réactiver cette note"
            End If
            
            If Memo_Aut.Xspécial Then arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "41 - Effacer cette note"
End Select

For K = 1 To 9
    For I = 1 To arrSelect_Update_Nb
        If Mid$(arrSelect_Update(I), 1, 1) = K Then
            arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = K & "0 --------------------------------------"
            Exit For
        End If
    Next I
Next K

If Not blnROPINFIDTL_Ok Then
    If Not blnYROPINF0_12X_Aut Then
        cmdSelect_Update.Enabled = False
    Else
        arrSelect_Update_Nb = 1: arrSelect_Update(arrSelect_Update_Nb) = "12X- Réaffecter(Resp,Ech) cette action"
        arrSelect_Update_Nb = arrSelect_Update_Nb + 1: arrSelect_Update(arrSelect_Update_Nb) = "02I- Insérer une action avant cette action"
    End If
End If

'cmdSelect_Update.ListIndex = 0

fraSelect_Update_B.Enabled = False
fraDétail_Update_B.Enabled = False
cmdSelect_Update_Ok.Visible = False
cmdSelect_Update_Close.Visible = False
cmdSelect_Update_Display (arrSelect_Update(1))
cmdSelect_Update.Height = 250 * cmdSelect_Update.ListCount
cmdSelect_Update.Visible = True


' à revoir
'cmdSelect_Update_Display (arrSelect_Update(arrSelect_Update_K))
'If arrSelect_Update_K <> 1 Then
'    blnSelect_Update_EnCours = False
'    cmdSelect_Update.ListIndex = arrSelect_Update_K
'End If

End Sub

Public Sub tvwSelect_STAK_Processus(kProcessus As Long)

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
If blnControl Then
    chkUpdate_ROPINFMAIL_D.Caption = tvwSelect_Display_USR(txtUpdate_ROPDOSGUSR)
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
lblUpdate_ROPDOSIAMJ.BackColor = &HC0F0FF

End Sub


Private Sub txtUpdate_ROPDOSIREF_GotFocus()
txt_GotFocus txtUpdate_ROPDOSIREF

End Sub


Private Sub txtUpdate_ROPDOSIREF_LostFocus()
txt_LostFocus txtUpdate_ROPDOSIREF

End Sub


Private Sub txtUpdate_ROPDOSIUSR_Click()
If blnControl Then
    chkUpdate_ROPINFMAIL_I.Caption = tvwSelect_Display_USR(txtUpdate_ROPDOSIUSR)
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

lblUpdate_ROPINFGECH.BackColor = focusUsr.BackColor
End Sub

Private Sub txtUpdate_ROPINFGECH_LostFocus()
If cmdSelect_SQL_K = "2" _
And txtUpdate_ROPINFGECH_Old <> txtUpdate_ROPINFGECH Then blntxtUpdate_ROPINFGECH_Change = True
lblUpdate_ROPINFGECH.BackColor = &HC0F0FF

End Sub

Private Sub txtUpdate_ROPINFGTXT_Click()
'If fraSelect_Update_B.Enabled Then
If Not txtUpdate_ROPINFGTXT.Locked Then
    chkSelect_Update_B.Value = "0"
    libUpdate_ROPINFGTXT.Visible = True
End If
End Sub


Private Sub txtUpdate_ROPINFGTXT_GotFocus()
txt_GotFocus txtUpdate_ROPINFGTXT

End Sub


Private Sub txtUpdate_ROPINFGTXT_LostFocus()
txt_LostFocus txtUpdate_ROPINFGTXT
Select Case cmdSelect_SQL_K
    Case "0", "2", "2M":  chkSelect_Update_B.Value = "1"
'    Case Else: libUpdate_ROPINFGTXT.Visible = False
End Select
End Sub


Private Sub txtUpdate_ROPINFGUO_GotFocus()
txt_GotFocus txtUpdate_ROPINFGUO

End Sub

Private Sub txtUpdate_ROPINFGUO_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtUpdate_ROPINFGUO)
End Sub



Public Sub cmdSelect_Update_Ok_04_ISRV()
Dim X As String, K As Integer
If IsNull(sqlYBIATAB0_Read("ROPDOSGUSR", newYROPDOS0.ROPDOSIUSR, "", X)) Then
    newYROPDOS0.ROPDOSISRV = "_" & Mid$(X, 26, 3)
Else
    newYROPDOS0.ROPDOSISRV = ""
End If
Call cmdSelect_Update_Ok_GSRV(newYROPDOS0.ROPDOSIUSR, newYROPDOS0.ROPDOSISRV)
End Sub

Public Sub cmdSelect_Update_Ok_GSRV(lUSR As String, lSRV As String)
Dim X As String, K As Integer
If Mid$(lUSR, 1, 2) = "_S" Then
    lSRV = lUSR
Else
    If IsNull(sqlYBIATAB0_Read("ROPDOSGUSR", lUSR, "", X)) Then
        lSRV = "_" & Mid$(X, 26, 3)
    Else
        lSRV = ""
    End If
End If

End Sub

Public Sub cmdPrint0_Dossier()
Dim mROPINFIDP As Long
Dim V, K As Long, X As String


prtROPDOS_Dossier oldYROPDOS0
mROPINFIDP = 1
For K = 1 To arrYROPINF0_Nb
    prtROPDOS_Détail_1 arrYROPINF0(K)
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
Dim blnOpen As Boolean, blnprtROPDOS_Détail_2P As Boolean
Dim xSql As String

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
    
    blnprtROPDOS_Détail_2P = False
    If mYROPPRT0.ROPPRTDEST <> xYROPPRT0.ROPPRTDEST Then
        If mYROPPRT0.ROPPRTDEST <> "" Then cmdPrint0_Echéancier_Destinataire
        blnOpen = True
        blnprtROPDOS_Détail_2P = True
        prtROPDOS_Open 2, "Risques Opérationnels : Echéancier " & xYROPPRT0.ROPPRTDEST & " au " & DSys
    End If
    
    If mYROPPRT0.ROPPRTID <> xYROPPRT0.ROPPRTID _
    Or mYROPPRT0.ROPPRTIDP <> xYROPPRT0.ROPPRTIDP Then
        blnprtROPDOS_Détail_2P = True
        xSql = "select *  from " & paramIBM_Library_SABSPE & ".YROPINF0 " _
         & " where ROPINFID = " & xYROPPRT0.ROPPRTID _
         & " and ROPINFIDP = " & xYROPPRT0.ROPPRTIDP
        Set rsSab = cnsab.Execute(xSql)
        If rsSab.EOF Then
            Call rsYROPINF0_Init(xYROPINF0)
        Else
            V = rsYROPINF0_GetBuffer(rsSab, xYROPINF0)
            If Not IsNull(V) Then Call rsYROPINF0_Init(xYROPINF0)
        End If
    End If
      
   ' If blnprtROPDOS_Détail_2P Then prtROPDOS_Détail_2P arrYROPDOS0(xYROPPRT0.ROPPRTarrD), xYROPINF0
   ' Call prtROPDOS_Détail_2A(arrYROPINF0(xYROPPRT0.ROPPRTarrI), blnprtROPDOS_Détail_2P)
   
    If blnprtROPDOS_Détail_2P Then
        prtROPDOS_Dossier arrYROPDOS0(xYROPPRT0.ROPPRTarrD)
        Call prtROPDOS_Détail_2(xYROPINF0)
    End If
    Call prtROPDOS_Détail_2(arrYROPINF0(xYROPPRT0.ROPPRTarrI))
   
    mYROPPRT0 = xYROPPRT0

Next K


If blnOpen Then cmdPrint0_Echéancier_Destinataire


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

End Sub
Public Sub fraAut_Reset()

On Error Resume Next

fraAut_Update.Visible = False
rsYBIATAB0_Init oldParam
Call lstZMNURUT0_Load_Actif_Production(lstAut_Usr)

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
lstAut_ROPDOSGUSR_Display (wIndex)

End Sub

Public Sub fraParam_Display()
fraParam_Update.Visible = True

txtParam_BIATABK1 = Trim(oldParam.BIATABK1): txtParam_BIATABK1.Enabled = False
txtParam_BIATABK2 = Trim(oldParam.BIATABK2): txtParam_BIATABK2.Enabled = False
txtParam_BIATABTXT = Trim(oldParam.BIATABTXT): txtParam_BIATABTXT.Enabled = False

End Sub
Public Function fraParam_Update_Control()
Dim blnUpdate_Control As Boolean
Dim X As String

blnUpdate_Control = True
newParam = oldParam
newParam.BIATABK1 = Trim(txtParam_BIATABK1)
newParam.BIATABK2 = Trim(txtParam_BIATABK2)
X = Trim(txtParam_BIATABTXT)
If X = "" Then
    blnUpdate_Control = False
    Call lstErr_AddItem(lstErr, cmdContext, "?_________préciser le libellé")
End If

newParam.BIATABTXT = X

If blnUpdate_Control Then
    fraParam_Update_Control = Null
Else
    fraParam_Update_Control = "<Fin du contrôle des données "
End If
End Function


Public Sub Form_Init_RODOSGUSR()
Dim X As String, xK1 As String, xUsr As String, xHabilitation As String
Dim K As Integer

For K = 0 To 100
    arrROPDOSISRV_Mail(K) = ""
Next K

blnROPDOSQUAL = False
txtUpdate_ROPDOSGUSR.Clear
txtUpdate_ROPDOSGUSR.AddItem "?"
txtUpdate_ROPDOSIUSR.Clear
txtUpdate_ROPDOSIUSR.AddItem "?"
txtUpdate_ROPINFGUSR.Clear
txtUpdate_ROPINFGUSR.AddItem "?"
lstUpdate_ROPINFMAIL.Clear

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'ROPDOSGUSR'"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    xK1 = rsSab("BIATABK1")
    xUsr = Trim(xK1)
    X = rsSab("BIATABTXT")
    If Mid$(X, 26, 1) = "S" Then
        K = Val(Mid$(X, 27, 2))
        xHabilitation = Mid$(X, 28 + K, 1)
        If xHabilitation = "R" Or xHabilitation = "D" Then arrROPDOSISRV_Mail(K) = arrROPDOSISRV_Mail(K) & xK1
        If Mid$(X, 1, 1) = "D" Then txtUpdate_ROPDOSGUSR.AddItem xUsr
        If Mid$(X, 2, 1) = "A" Then txtUpdate_ROPINFGUSR.AddItem xUsr: txtUpdate_ROPDOSIUSR.AddItem xUsr
        If xUsr = usrName_UCase Then
            currentROPDOSISRV = "_" & Mid$(X, 26, 3)
            currentROPDOSISRV_Nom = Trim(tvwSelect_Display_USR(currentROPDOSISRV))
            currentROPDOSISRV_Hab = Mid$(X, 29, 99)
            currentROPDOSISRV_Rôle = xHabilitation
            If Mid$(X, 3, 1) = "P" Then
                fraParam.Visible = True
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
            
        End If
        lstUpdate_ROPINFMAIL.AddItem Mid$(xUsr, 1, 10)
    End If
    rsSab.MoveNext
Loop
'________________________________________________________________________________
X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'ROPDOSISRV' order by BIATABTXT"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    xUsr = "_" & Format(rsSab("BIATABK1"), "          ") & "_" & Trim(Mid$(rsSab("BIATABTXT"), 1, 12))
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
X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'ROPDOSISRV' order by BIATABK1"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    xK1 = Trim(rsSab("BIATABK1"))
    K = Val(Mid$(xK1, 2, 2))
    arrROPDOSISRV_K1(K) = xK1
    X = rsSab("BIATABTXT")
    arrROPDOSISRV_Code(K) = Mid$(X, 1, 12)
    arrROPDOSISRV_Lib(K) = Mid$(X, 13, 64)
    lstAut_ROPDOSISRV.AddItem xK1 & " " & arrROPDOSISRV_Code(K) & " " & arrROPDOSISRV_Lib(K)
    arrROPDOSISRV_ListIndex(K) = lstAut_ROPDOSISRV.ListCount
    rsSab.MoveNext
Loop

End Sub

Public Sub cmdSelect_Update_Display(lItem As String)
Dim K As Integer
cmdSelect_Update.Clear
For K = 1 To arrSelect_Update_Nb
    If lItem = arrSelect_Update(K) Then
        cmdSelect_Update.AddItem " >" & arrSelect_Update(K)
    Else
        cmdSelect_Update.AddItem "- " & arrSelect_Update(K)
    End If
    
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
Dim xSql As String
Dim Xdisplay As String * 120

Set mSelect_Node = Nothing
tvwSelect.Nodes.Clear
For K = 1 To selYROPDOS0_Nb
    xYROPDOS0 = selYROPDOS0(K)
    blnOk = False
    If xYROPDOS0.ROPDOSGPRV = "U" Then
        If Trim(xYROPDOS0.ROPDOSGUSR) = usrName_UCase Then blnOk = True
    Else
        If cmdSelect_SQL_K <> "1X" Then
            blnOk = True
        Else
            Select Case xYROPDOS0.ROPDOSGPRV
                Case "W": blnOk = True
                Case Else:
                    If cmdSelect_SQL_1_Hab(Trim(xYROPDOS0.ROPDOSGUSR)) Then
                        blnOk = True
                    Else
                        If cmdSelect_SQL_1_Hab(Trim(xYROPDOS0.ROPDOSIUSR)) Then blnOk = True
                    End If
            End Select
        End If
    End If
    
    If blnOk Then
        X = "D" & Format$(xYROPDOS0.ROPDOSID, "000000000")
         If xYROPDOS0.ROPDOSID < 1000 Then
            
            Set xNode = tvwSelect.Nodes.Add(, , X, Format$(xYROPDOS0.ROPDOSID, "0000") & " - " & tvwSelect_ROPDOSXAPP_Libellé(xYROPDOS0.ROPDOSXDOM, xYROPDOS0.ROPDOSXAPP))
        Else
            Set xNode = tvwSelect.Nodes.Add(, , X, Format$(xYROPDOS0.ROPDOSID, "0000") & " - " & xYROPDOS0.ROPDOSXDOM & " " & xYROPDOS0.ROPDOSXAPP & " (" & tvwSelect_Display_USR(xYROPDOS0.ROPDOSGUSR) & " | " & tvwSelect_Display_USR(xYROPDOS0.ROPDOSIUSR) & ")")
        End If
        
        tvwSelect_YROPDOS0_Forecolor
        
        xSql = "select ROPINFGTXT from " & paramIBM_Library_SABSPE & ".YROPINF0" _
            & " where ROPINFID = " & xYROPDOS0.ROPDOSID _
            & " and ROPINFIDP = 1 and ROPINFIDT = 0 and ROPINFIDT2= 1"
     
        Set rsSab = cnsab.Execute(xSql)
        
        If Not rsSab.EOF Then
        Xdisplay = rsSab("ROPINFGTXT")
            Set xNode = tvwSelect.Nodes.Add(, , X & "+", Format$(xYROPDOS0.ROPDOSID, "0000") & " : " & xYROPDOS0.ROPDOSXID & " > " & Xdisplay)
            xNode.ForeColor = &H808080 'vbBlack '&H202020
            xNode.BackColor = &HE0FFFF  ' &HE0E0E0
        End If
    End If
Next K
tvwSelect.Visible = True
fraSelect.Visible = True
cmdPrint.Enabled = True


End Sub

Public Sub cmdSelect_Update_Init_99()
Me.Enabled = False: Me.MousePointer = vbHourglass

prtROPDOS_Open 1, "Risques Opérationnels"
cmdPrint0_Dossier
prtROPDOS_Close 1


Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdSelect_Update_Init_98()
lstUpdate_ROPINFMAIL.TopIndex = 0
lstUpdate_ROPINFMAIL.Visible = True
cmdSelect_Update_Ok.Visible = True
fraUpdate_ROPINFMAIL.Visible = False
cmdSelect_Update_Ok.Caption = "Envoyer"

End Sub
Public Sub cmdSelect_Update_Ok_98()
Me.Enabled = False: Me.MousePointer = vbHourglass

mailYROPDOS0 = oldYROPDOS0: mailYROPINF0 = arrYROPINF0(1)

Call cmdSendMail("X")
cmdSelect_Update_Reset
Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdPrint0_Dossier_All()
Dim K As Integer, xSql As String
On Error Resume Next

prtROPDOS_Open 1, "Risques Opérationnels"

For K = 1 To selYROPDOS0_Nb
    
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 where ROPDOSID =" & selYROPDOS0(K).ROPDOSID
    Set rsSab = cnsab.Execute(xSql)
    
    If Not rsSab.EOF Then
        V = rsYROPDOS0_GetBuffer(rsSab, oldYROPDOS0)
        arrYROPINF0_SQL oldYROPDOS0.ROPDOSID
    End If
    tvwSelect_STAK
    cmdPrint0_Dossier
Next K
prtROPDOS_Close 1
fraSelect_Update.Visible = False
End Sub

Public Function cmdSendMail_Txt(lTxt As String, lK As String) As String
Dim lenX As Long, K As Integer, K1 As Integer
Dim I As Integer
Dim htmlTxt As String, blnEnd As Boolean
Dim wNb As Integer, wReturn As String

If lK = "H" Then
    wNb = 150
    wReturn = "<BR>"
    htmlTxt = vbCrLf
Else
    wNb = 90 '65
    wReturn = vbCrLf
    htmlTxt = ""
End If

htmlTxt = ""
K = 1
lenX = Len(lTxt)
blnEnd = False
Do
    K1 = InStr(K, lTxt, vbCrLf)
    If K1 > 0 Then
        For I = K To K1 Step wNb
            If I + wNb < K1 Then
                htmlTxt = htmlTxt & Mid$(lTxt, I, wNb) & wReturn
            Else
                htmlTxt = htmlTxt & Mid$(lTxt, I, K1 - I) & wReturn
            End If
        Next I
        K = K1 + 2
    Else
        blnEnd = True
         For I = K To lenX Step wNb
            If I + wNb < lenX Then
                htmlTxt = htmlTxt & Mid$(lTxt, I, wNb) & wReturn
            Else
                htmlTxt = htmlTxt & Mid$(lTxt, I, lenX - I + 1)
            End If
        Next I
    End If
Loop Until blnEnd
cmdSendMail_Txt = htmlTxt
End Function

Public Function cmdSendMail_Sta(lSta As String, lSta_backColor As String) As String
'
'
Select Case lSta
    Case " ": cmdSendMail_Sta = "<Font color = #0000FF > " 'en cours"
    Case "F", "+", "$", "*": lSta_backColor = "bgcolor = #A0FFA0": cmdSendMail_Sta = wTD_Txt_ForeColor & "ok"
    Case "A", "-", "%", "£": lSta_backColor = "bgcolor = #FF0000": cmdSendMail_Sta = wTD_Txt_ForeColor & "Annulé"
    Case Else: lSta_backColor = "bgcolor = #FFFFFF": cmdSendMail_Sta = wTD_Txt_ForeColor & oldYROPDOS0.ROPDOSSTA
End Select
End Function

Public Function cmdSendMail_Ech(lEch As String, lSta As String) As String
If lSta <> " " Then
    cmdSendMail_Ech = wTD_Txt_ForeColor & dateImp10(lEch)
Else
    If lEch > DSys Then
        cmdSendMail_Ech = "<Font color = #0000FF>" & dateImp10(lEch)
    Else
        If lEch = DSys Then
            cmdSendMail_Ech = "<Font color = #FF00FF>" & dateImp10(lEch)
        Else
            cmdSendMail_Ech = "<Font color = #FF0000>" & dateImp10(lEch)
        End If
    End If
End If
End Function


Public Sub lstAut_ROPDOSGUSR_Display(lIndex As Long)
Dim K As Integer, K2 As Integer, X As String

lstAut_ROPDOSGUSR.Clear
lIndex = arrROPDOSISRV_ListIndex(Val(Mid$(newAut.BIATABTXT, 27, 2)))
For K = 1 To 99
    X = "S" & Format(K, "00")
    If Mid$(newAut.BIATABTXT, K + 28, 1) <> " " Then
        lstAut_ROPDOSGUSR.AddItem Mid$(newAut.BIATABTXT, K + 28, 1) & " - " & arrROPDOSISRV_Code(K)
    End If
Next K

End Sub

Public Function tvwSelect_ROPDOSXAPP_Libellé(lROPDOSXDOM As String, lROPDOSXAPP As String) As String
Dim X As String
If IsNull(sqlYBIATAB0_Read("ROPDOSXAPP", lROPDOSXDOM, lROPDOSXAPP, X)) Then
    tvwSelect_ROPDOSXAPP_Libellé = X
Else
    tvwSelect_ROPDOSXAPP_Libellé = lROPDOSXDOM & "-" & lROPDOSXAPP
End If


End Function

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

Public Sub cmdSelect_Update_Init_24_Ok()

mROPINFSTA_Value = "F": mROPINFSTA_Set = "$": mROPINFSTA_Where = " ": mROPINFSTAK_Set = "V"
newYROPDOS0.ROPDOSSTA = mROPINFSTA_Value
newYROPDOS0.ROPDOSSTAK = "V"
cmdSelect_Update_Ok_Transaction "Update"
blnmSelect_Node = True
tvwSelect_NodeClick tvwSelect.Nodes("D" & Format$(oldYROPDOS0.ROPDOSID, "000000000"))

End Sub

Public Sub cmdSelect_Update_Init_23_Ok()
    
mROPINFSTA_Value = "F": mROPINFSTA_Set = "*": mROPINFSTA_Where = " ": mROPINFSTAK_Set = "V"
cmdDétail_Update_Ok
If xYROPDOS0.ROPDOSSTAK = "V" Then cmdSelect_Update_Init_24_Ok ' fermeture dossier

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
    Case "Z": optAut_ROPDOSISRV_X = True
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
Dim xSql As String
Me.Enabled = False

oldYROPDOS0.ROPDOSID = 10
xSql = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0 where ROPDOSID = 10"
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then
    blntxtUpdate_ROPINFGECH_Change = False
    V = rsYROPDOS0_GetBuffer(rsSab, oldYROPDOS0)
    arrYROPINF0_SQL oldYROPDOS0.ROPDOSID
   cmdSelect_SQL_K = "2"
    cmdSelect_SQL_2_Duplication
End If
Me.Enabled = True

End Sub


Public Function cmdSelect_Update_Init_00_Hab(lGUSR As String) As Boolean
cmdSelect_Update_Init_00_Hab = False
If usrName_UCase = lGUSR Then
    cmdSelect_Update_Init_00_Hab = True
Else
    If Mid$(lGUSR, 1, 2) = "_S" Then
        X = Mid$(currentROPDOSISRV_Hab, Val(Mid$(lGUSR, 3, 2)), 1)
        If X = "H" Or X = "R" Or X = "D" Then cmdSelect_Update_Init_00_Hab = True
    End If
End If

End Function
Public Function cmdSelect_SQL_1_Hab(lGUSR As String) As Boolean
cmdSelect_SQL_1_Hab = False
If usrName_UCase = lGUSR Then
    cmdSelect_SQL_1_Hab = True
Else
    If Mid$(lGUSR, 1, 2) = "_S" Then
        X = Mid$(currentROPDOSISRV_Hab, Val(Mid$(lGUSR, 3, 2)), 1)
        If X = "H" Or X = "R" Or X = "D" Or X = "C" Or X = "I" Then cmdSelect_SQL_1_Hab = True
    End If
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
        Case "P": chkUpdate_ROPINFMAIL_P.Caption = tvwSelect_Display_USR(txtUpdate_ROPINFGUSR)
        Case "A", "F": chkUpdate_ROPINFMAIL_A.Caption = tvwSelect_Display_USR(txtUpdate_ROPINFGUSR)
    
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



Public Sub fraSelect_Display_Reset()
lblUpdate_ROPDOSGECH.BackColor = &HC0F0FF
txtUpdate_ROPDOSGECH.ToolTipText = "Echéance à laquelle le dossier doit être clôturer"
lblUpdate_ROPDOSIAMJ.BackColor = &HC0F0FF
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
lblUpdate_ROPINFGECH.BackColor = &HC0F0FF
txtUpdate_ROPINFGECH.ToolTipText = "Echéance >= jour et =< échéance dossier"

End Sub

Public Sub cmdSelect_SQL_7_Destinataire()
Dim K1 As Integer, K2 As Integer
Dim xUsr As String, kLen As Integer

If xYROPDOS0.ROPDOSGPRV = "U" Then
    lstW.AddItem xYROPDOS0.ROPDOSIUSR & X
Else
    If Mid$(xYROPINF0.ROPINFGUSR, 1, 1) <> "_" Then
        lstW.AddItem xYROPINF0.ROPINFGUSR & X
    Else
        K1 = Val(Mid$(xYROPINF0.ROPINFGUSR, 3, 2))
        xUsr = arrROPDOSISRV_Mail(K1)
        kLen = Len(xUsr)
        For K2 = 1 To kLen Step 12
            lstW.AddItem Mid$(xUsr, K2, 12) & X
        Next K2

    End If
    
End If

End Sub

Public Sub cmdPrint0_Echéancier_Destinataire()
XPrt.DrawWidth = 8
XPrt.CurrentY = XPrt.CurrentY + 100
XPrt.Line (prtMinX, XPrt.CurrentY + prtlineHeight)-(prtMaxX, XPrt.CurrentY + prtlineHeight), prtLineColor
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight + 20

prtROPDOS_Close 2

End Sub
