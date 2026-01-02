VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBDF_CMP 
   AutoRedraw      =   -1  'True
   Caption         =   "BDF_CMP : déclaration  cartographie des moyens de paiement"
   ClientHeight    =   10755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "BDF_CMP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10755
   ScaleWidth      =   13875
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   45
      Width           =   5055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10092
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   13812
      _ExtentX        =   24368
      _ExtentY        =   17806
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Sélection"
      TabPicture(0)   =   "BDF_CMP.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Statistiques"
      TabPicture(1)   =   "BDF_CMP.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstW"
      Tab(1).Control(1)=   "fgStatistiques"
      Tab(1).ControlCount=   2
      Begin VB.ListBox lstW 
         Height          =   255
         Left            =   -67800
         Sorted          =   -1  'True
         TabIndex        =   33
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame fraSelect 
         Height          =   9492
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13560
         Begin MSComCtl2.DTPicker txtSelect_AmjMax 
            Height          =   300
            Left            =   12240
            TabIndex        =   60
            Top             =   720
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
            Format          =   109510659
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin VB.ListBox lstSelect 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5310
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   37
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Frame fraUpdate 
            BackColor       =   &H00F0F0F0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   7932
            Left            =   8280
            TabIndex        =   11
            Top             =   1200
            Width           =   5175
            Begin VB.Frame fraUpdate_B 
               BackColor       =   &H00D0D0D0&
               Caption         =   "Grille - Ligne - Colonne"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2172
               Left            =   120
               TabIndex        =   52
               Top             =   5760
               Width           =   4935
               Begin VB.ComboBox cboUpdate_BDFCMP2008 
                  Height          =   288
                  Left            =   360
                  Style           =   2  'Dropdown List
                  TabIndex        =   62
                  Top             =   960
                  Width           =   4455
               End
               Begin VB.ComboBox cboUpdate_BDFCMPSTAT 
                  Height          =   288
                  Left            =   360
                  Style           =   2  'Dropdown List
                  TabIndex        =   59
                  Top             =   480
                  Width           =   4455
               End
               Begin VB.CommandButton cmdUpdate_Ok 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Enregistrer"
                  Height          =   648
                  Left            =   3000
                  Style           =   1  'Graphical
                  TabIndex        =   55
                  Top             =   1320
                  Width           =   1575
               End
               Begin VB.CommandButton cmdUpdate_Quit 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Abandonner"
                  Height          =   525
                  Left            =   1920
                  Style           =   1  'Graphical
                  TabIndex        =   54
                  Top             =   1440
                  Width           =   972
               End
               Begin VB.CommandButton cmdUpdate_Annuler 
                  BackColor       =   &H000000FF&
                  Caption         =   "Annuler/Reprendre"
                  Height          =   525
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   53
                  Top             =   1440
                  Width           =   1332
               End
            End
            Begin VB.Frame fraUpdate_A 
               BackColor       =   &H00F0F0F0&
               Height          =   5532
               Left            =   120
               TabIndex        =   12
               Top             =   120
               Width           =   4935
               Begin VB.TextBox txtUpdate_BDFCMP59PI 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   372
                  Left            =   3840
                  TabIndex        =   68
                  Text            =   "Text1"
                  Top             =   5040
                  Width           =   852
               End
               Begin VB.TextBox txtUpdate_BDFCMP50PI 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Left            =   1200
                  TabIndex        =   67
                  Text            =   "Text1"
                  Top             =   5160
                  Width           =   732
               End
               Begin VB.TextBox txtUpdate_BDFCMPMTk 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Left            =   1200
                  TabIndex        =   66
                  Text            =   "Text1"
                  Top             =   4800
                  Width           =   1812
               End
               Begin VB.TextBox txtUpdate_BDFCMPCREG 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   50
                  Top             =   1200
                  Width           =   615
               End
               Begin VB.TextBox txtUpdate_BDFCMPPAYS 
                  Height          =   285
                  Left            =   3960
                  TabIndex        =   48
                  Top             =   3960
                  Width           =   495
               End
               Begin VB.TextBox txtUpdate_BDFCMPUSR 
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   47
                  Text            =   "usr"
                  Top             =   4320
                  Width           =   1935
               End
               Begin VB.TextBox txtUpdate_BDFCMPXDBN 
                  Height          =   285
                  Left            =   3960
                  TabIndex        =   45
                  Top             =   3120
                  Width           =   615
               End
               Begin VB.TextBox txtUpdate_BDFCMPXCRN 
                  Height          =   285
                  Left            =   3960
                  TabIndex        =   44
                  Top             =   2280
                  Width           =   615
               End
               Begin VB.TextBox txtUpdate_BDFCMPMONE 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2640
                  TabIndex        =   41
                  Top             =   1200
                  Width           =   2055
               End
               Begin VB.TextBox txtUpdate_BDFCMPBBIC 
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   40
                  Top             =   3960
                  Width           =   2055
               End
               Begin VB.TextBox txtUpdate_BDFCMPSTA 
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   25
                  Top             =   4320
                  Width           =   615
               End
               Begin MSComCtl2.DTPicker txtUpdate_BDFCMPDCRE 
                  Height          =   300
                  Left            =   1200
                  TabIndex        =   24
                  Top             =   1800
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   16777215
                  CalendarForeColor=   0
                  CalendarTitleBackColor=   8421504
                  CalendarTitleForeColor=   16777215
                  CalendarTrailingForeColor=   12632256
                  CustomFormat    =   "dd  MM yyy"
                  Format          =   109510659
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin VB.TextBox txtUpdate_BDFCMPXCR 
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   22
                  Top             =   2280
                  Width           =   2055
               End
               Begin VB.TextBox txtUpdate_BDFCMPXDB 
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   21
                  Top             =   3120
                  Width           =   2055
               End
               Begin VB.TextBox txtUpdate_BDFCMPMON 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2640
                  TabIndex        =   19
                  Top             =   720
                  Width           =   2055
               End
               Begin VB.TextBox txtUpdate_BDFCMPDEV 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   18
                  Top             =   720
                  Width           =   615
               End
               Begin VB.TextBox txtUpdate_BDFCMPDOS 
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   16
                  Top             =   240
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_BDFCMPOPE 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   14
                  Top             =   240
                  Width           =   615
               End
               Begin VB.TextBox txtUpdate_BDFCMPNAT 
                  Height          =   285
                  Left            =   2160
                  TabIndex        =   13
                  Top             =   240
                  Width           =   615
               End
               Begin MSComCtl2.DTPicker txtUpdate_BDFCMPDOPE 
                  Height          =   300
                  Left            =   3240
                  TabIndex        =   42
                  Top             =   1800
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   16777215
                  CalendarForeColor=   0
                  CalendarTitleBackColor=   8421504
                  CalendarTitleForeColor=   16777215
                  CalendarTrailingForeColor=   12632256
                  CustomFormat    =   "dd  MM yyy"
                  Format          =   109510659
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin VB.Label lblUpdate_BDFCMP59PI 
                  Caption         =   "BEN"
                  Height          =   252
                  Left            =   3120
                  TabIndex        =   65
                  Top             =   5160
                  Width           =   612
               End
               Begin VB.Label lblUpdate_BDFCMP50PI 
                  Caption         =   "DO"
                  Height          =   252
                  Left            =   120
                  TabIndex        =   64
                  Top             =   5160
                  Width           =   732
               End
               Begin VB.Label Label2 
                  Caption         =   "Label2"
                  Height          =   252
                  Left            =   0
                  TabIndex        =   63
                  Top             =   0
                  Width           =   4092
               End
               Begin VB.Label lblUpdate_BDFCMPUSR 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "User"
                  Height          =   255
                  Left            =   1920
                  TabIndex        =   58
                  Top             =   4560
                  Width           =   735
               End
               Begin VB.Label Label1 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "EUR"
                  Height          =   255
                  Left            =   2040
                  TabIndex        =   57
                  Top             =   1200
                  Width           =   495
               End
               Begin VB.Label lblUpdate_BDFCMPPAYS 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "Pays"
                  Height          =   255
                  Left            =   3360
                  TabIndex        =   56
                  Top             =   4080
                  Width           =   615
               End
               Begin VB.Label lblUpdate_BDFCMPCREG 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "Code rég"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   51
                  Top             =   1200
                  Width           =   975
               End
               Begin VB.Label lblUpdate_BDFCMPSTA 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "Statut"
                  Height          =   252
                  Left            =   120
                  TabIndex        =   49
                  Top             =   4320
                  Width           =   972
               End
               Begin VB.Label libUpdate_BDFCMPXDB 
                  BackColor       =   &H00D0D0D0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "x"
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
                  Left            =   120
                  TabIndex        =   46
                  Top             =   3600
                  Width           =   4455
               End
               Begin VB.Label lblUpdate_BDFCMPXCR 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "Crédit"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   43
                  Top             =   2280
                  Width           =   975
               End
               Begin VB.Label lblUpdate_BDFCMPBBIC 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "BIC bénéficiaire "
                  Height          =   252
                  Left            =   120
                  TabIndex        =   38
                  Top             =   3960
                  Width           =   1212
               End
               Begin VB.Label lblUpdate_BDFCMPDCRE 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "Créé / D Opé"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   26
                  Top             =   1800
                  Width           =   975
               End
               Begin VB.Label libUpdate_BDFCMPXCR 
                  BackColor       =   &H00D0D0D0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "x"
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
                  Left            =   120
                  TabIndex        =   23
                  Top             =   2640
                  Width           =   4575
               End
               Begin VB.Label lblUpdate_BDFCMPXDB 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "Débit"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   20
                  Top             =   3120
                  Width           =   975
               End
               Begin VB.Label lblUpdate_BDFCMPMON 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "Montant"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   17
                  Top             =   840
                  Width           =   975
               End
               Begin VB.Label lblUpdate_BDFCMPDOS 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "Dossier"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF80FF&
                  Height          =   495
                  Left            =   120
                  TabIndex        =   15
                  Top             =   240
                  Width           =   1095
               End
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   8268
            Left            =   120
            TabIndex        =   8
            Top             =   1080
            Width           =   13440
            _ExtentX        =   23707
            _ExtentY        =   14579
            _Version        =   393216
            Rows            =   1
            Cols            =   12
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
            FormatString    =   $"BDF_CMP.frx":0044
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
         Begin VB.ComboBox cboSelect_SQL 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   9
            Text            =   "cboSelect_SQL"
            Top             =   260
            Width           =   3015
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Rechercher"
            Height          =   525
            Left            =   11160
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   1815
         End
         Begin VB.Frame fraSelect_Options_1 
            Height          =   1005
            Left            =   3240
            TabIndex        =   6
            Top             =   120
            Width           =   7035
            Begin VB.CheckBox chkSelect_BDFCMPSTA 
               Alignment       =   1  'Right Justify
               Caption         =   "Inclure 'Ann' 'I...'"
               Height          =   255
               Left            =   5160
               TabIndex        =   61
               Top             =   700
               Width           =   1575
            End
            Begin VB.TextBox txtSelect_BDFCMPSTAT 
               Height          =   285
               Left            =   6120
               TabIndex        =   36
               Top             =   240
               Width           =   615
            End
            Begin VB.CheckBox chkSelect_BDFCMPDCRE 
               Caption         =   "Période de création"
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox txtSelect_BDFCMPNAT 
               Height          =   285
               Left            =   4440
               TabIndex        =   29
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox txtSelect_BDFCMPOPE 
               Height          =   285
               Left            =   4440
               TabIndex        =   28
               Top             =   240
               Width           =   615
            End
            Begin MSComCtl2.DTPicker txtSelect_BDFCMPDCRE 
               Height          =   300
               Left            =   2040
               TabIndex        =   27
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   109510659
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_BDFCMPDCRE_Max 
               Height          =   300
               Left            =   2040
               TabIndex        =   34
               Top             =   600
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   109510659
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_BDFCMPSTAT 
               Caption         =   "Code    BDF CMP"
               Height          =   375
               Left            =   5160
               TabIndex        =   35
               Top             =   240
               Width           =   855
            End
            Begin VB.Label lblSelect_BDFCMPNAT 
               Caption         =   "Nature"
               Height          =   255
               Left            =   3600
               TabIndex        =   31
               Top             =   720
               Width           =   855
            End
            Begin VB.Label lblSelect_BDFCMPOPE 
               Caption         =   "Code opération"
               Height          =   375
               Left            =   3600
               TabIndex        =   30
               Top             =   240
               Width           =   735
            End
         End
         Begin MSComCtl2.DTPicker txtSelect_AmjMin 
            Height          =   300
            Left            =   10680
            TabIndex        =   10
            Top             =   720
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
            Format          =   109510659
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgStatistiques 
         Height          =   8145
         Left            =   -74880
         TabIndex        =   39
         Top             =   600
         Width           =   13440
         _ExtentX        =   23707
         _ExtentY        =   14367
         _Version        =   393216
         Rows            =   1
         Cols            =   17
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
         FormatString    =   $"BDF_CMP.frx":00D0
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
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "BDF_CMP.frx":0183
      Style           =   1  'Graphical
      TabIndex        =   1
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
   Begin VB.Menu mnuselect 
      Caption         =   "mnuSelect"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Quit 
         Caption         =   "Abandonner"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint0_All 
         Caption         =   "Imprimer TOUS les courriers"
      End
   End
End
Attribute VB_Name = "frmBDF_CMP"
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
Dim intReturn As Integer
Dim BIA_BDFCMP_Aut As typeAuthorization
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

Dim mCLIENARA1 As String

'______________________________________________________________________

Dim xYBDFCMP0 As typeYBDFCMP0, meYBDFCMP0 As typeYBDFCMP0
Dim newYBDFCMP0 As typeYBDFCMP0, oldYBDFCMP0 As typeYBDFCMP0
Dim arrYBDFCMP0() As typeYBDFCMP0, arrYBDFCMP0_Nb As Long, arrYBDFCMP0_Max As Long, arrYBDFCMP0_Index As Long
Dim selYBDFCMP0() As typeYBDFCMP0, selYBDFCMP0_Nb As Long, selYBDFCMP0_Max As Long, selYBDFCMP0_Index As Long
Dim xZCLIENA0 As typeZCLIENA0

Dim xYBIAMVT0 As typeYBIAMVT0
Dim cmdSelect_Ok_Caption As String
Dim cmdSelect_SQL_K As String

Dim curDB As Currency, curCR As Currency
Dim selZCHGOPE0() As typeZCHGOPE0, selZCHGOPE0_Nb As Long, selZCHGOPE0_Max As Long, selZCHGOPE0_Index As Long
Dim xZCHGOPE0 As typeZCHGOPE0

Dim rsSabX As New ADODB.Recordset

Dim arrBDFCMPMON() As Currency
Dim wMM As Integer, wAAAA As Integer, arrBDFCMPMON_Dev As Integer

Dim fgStatistiques_FormatString As String, fgStatistiques_K As Integer
Dim fgStatistiques_RowDisplay As Integer, fgStatistiques_RowClick As Integer, fgStatistiques_ColClick As Integer
Dim fgStatistiques_ColorClick As Long, fgStatistiques_ColorDisplay As Long
Dim fgStatistiques_Sort1 As Integer, fgStatistiques_Sort2 As Integer
Dim fgStatistiques_SortAD As Integer, fgStatistiques_Sort1_Old As Integer
Dim fgStatistiques_arrIndex As Integer
Dim blnfgStatistiques_DisplayLine As Boolean

Dim meCV1 As typeCV, meCV2 As typeCV

Dim arrPAYS_UE(50) As String, arrPAYS_UE_Nb As Integer
'______________________________________________________________________

Dim mENCREGCL5 As String, mENCREGBC5 As String, mENCREGLC5 As String
Dim mENCREGCL5_D As String, mENCREGBC5_D As String, mENCREGLC5_D As String
Dim mENCREGCL5_C As String, mENCREGBC5_C As String, mENCREGLC5_C As String


Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel

Public Sub YBDFCMP0_V2008_Export()
On Error GoTo Error_Handler
Dim Nb As Long, wId As Long
Dim wFile As String, xSql As String
Dim nbk As Long, mtk As Currency
Dim X As String

rsYBDFCMP0_Init oldYBDFCMP0
wFile = "C:\temp\YBDFCMP0_V2008.xlsx"
X = MsgBox("export du fichier : " & wFile & " ?", vbYesNo, "YBDFCMP0_V2008")
If X <> vbYes Then Exit Sub

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
Set wsExcel = wbExcel.ActiveSheet
'__________________________________________________________________________________

Nb = 1: nbk = 0
Call lstErr_AddItem(lstErr, cmdContext, "Export en cours : " & Nb & " enregistrements"): DoEvents

wsExcel.Cells(Nb, 1) = "2008"
wsExcel.Cells(Nb, 2) = "MT"
wsExcel.Cells(Nb, 3) = "Route"
wsExcel.Cells(Nb, 4) = "CR nat"
wsExcel.Cells(Nb, 5) = "DB nat"
wsExcel.Cells(Nb, 6) = "C reg"
wsExcel.Cells(Nb, 7) = "Ser"
wsExcel.Cells(Nb, 8) = "Sse"
wsExcel.Cells(Nb, 9) = "Ope"
wsExcel.Cells(Nb, 10) = "Nat"
wsExcel.Cells(Nb, 11) = "N° dossier"
wsExcel.Cells(Nb, 12) = "Nb"
wsExcel.Cells(Nb, 13) = "MT €"


xSql = "select * from " & paramIBM_Library_SABSPE & ".YBDFCMP0" _
     & " order by BDFCMP2008, BDFCMPMTK, BDFCMPROUT, BDFCMPXCRN, BDFCMPXDBN, BDFCMPCREG, BDFCMPSER, BDFCMPSSE, BDFCMPOPE, BDFCMPNAT"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYBDFCMP0_GetBuffer(rsSab, xYBDFCMP0)
    If oldYBDFCMP0.BDFCMP2008 = xYBDFCMP0.BDFCMP2008 _
    And oldYBDFCMP0.BDFCMPMTK = xYBDFCMP0.BDFCMPMTK _
    And oldYBDFCMP0.BDFCMPROUT = xYBDFCMP0.BDFCMPROUT _
    And oldYBDFCMP0.BDFCMPXCRN = xYBDFCMP0.BDFCMPXCRN _
    And oldYBDFCMP0.BDFCMPXDBN = xYBDFCMP0.BDFCMPXDBN _
    And oldYBDFCMP0.BDFCMPCREG = xYBDFCMP0.BDFCMPCREG _
    And oldYBDFCMP0.BDFCMPSER = xYBDFCMP0.BDFCMPSER _
    And oldYBDFCMP0.BDFCMPSSE = xYBDFCMP0.BDFCMPSSE _
    And oldYBDFCMP0.BDFCMPOPE = xYBDFCMP0.BDFCMPOPE _
    And oldYBDFCMP0.BDFCMPNAT = xYBDFCMP0.BDFCMPNAT Then
        nbk = nbk + 1
        mtk = mtk + xYBDFCMP0.BDFCMPMONE
    Else
        Call YBDFCMP0_V2008_Export_Add(Nb, nbk, mtk)
        nbk = 1
        mtk = xYBDFCMP0.BDFCMPMONE
        oldYBDFCMP0 = xYBDFCMP0
    End If
    

    rsSab.MoveNext
Loop
Call YBDFCMP0_V2008_Export_Add(Nb, nbk, mtk)
Set rsSab = Nothing


'____________________________________________________________________________________
wbExcel.SaveAs wFile

wbExcel.Close
appExcel.Quit

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
Call lstErr_AddItem(lstErr, cmdContext, "Export terminé : " & Nb & " enregistrements"): DoEvents

'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub

Public Sub YBDFCMP0_Export()
On Error GoTo Error_Handler
Dim Nb As Long, wId As Long
Dim wFile As String, xSql As String
Dim nbk As Long
Dim X As String
    Dim iBackColor  As Integer

rsYBDFCMP0_Init oldYBDFCMP0
wFile = "C:\temp\YBDFCMP0.xlsx"
X = MsgBox("export du fichier : " & wFile & " ?", vbYesNo, "YBDFCMP0_V2008")
If X <> vbYes Then Exit Sub

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
Set wsExcel = wbExcel.ActiveSheet
'__________________________________________________________________________________

Nb = 1: nbk = 0
Call lstErr_AddItem(lstErr, cmdContext, "Export en cours : " & Nb & " enregistrements"): DoEvents

wsExcel.Cells(Nb, 1) = "SER"
wsExcel.Cells(Nb, 2) = "SSE"
wsExcel.Cells(Nb, 3) = "OPE"
wsExcel.Cells(Nb, 4) = "NAT"
wsExcel.Cells(Nb, 5) = "DOS"
wsExcel.Cells(Nb, 6) = "MON"
wsExcel.Cells(Nb, 7) = "DEVr"
wsExcel.Cells(Nb, 8) = "MONE"
wsExcel.Cells(Nb, 9) = "DCRE"
wsExcel.Cells(Nb, 10) = "DOPE"
wsExcel.Cells(Nb, 11) = "CREG"
wsExcel.Cells(Nb, 12) = "XDB"
wsExcel.Cells(Nb, 13) = "XDBN"
wsExcel.Cells(Nb, 14) = "XCR"
wsExcel.Cells(Nb, 15) = "XCRN"
wsExcel.Cells(Nb, 16) = "BBIC"
wsExcel.Cells(Nb, 17) = "PAYS"
wsExcel.Cells(Nb, 18) = "STAT"
wsExcel.Cells(Nb, 19) = "STA"
wsExcel.Cells(Nb, 20) = "UPDS"
wsExcel.Cells(Nb, 21) = "USR"
wsExcel.Cells(Nb, 22) = "SEQ"
wsExcel.Cells(Nb, 23) = "2008"
wsExcel.Cells(Nb, 24) = "SABK"
wsExcel.Cells(Nb, 25) = "MTK"
wsExcel.Cells(Nb, 26) = "ROUT"
wsExcel.Cells(Nb, 27) = "50PI"
wsExcel.Cells(Nb, 28) = "59PI"



xSql = "select * from " & paramIBM_Library_SABSPE & ".YBDFCMP0" _
     & " order by BDFCMP2008, BDFCMPMTK, BDFCMPROUT, BDFCMPXCRN, BDFCMPXDBN, BDFCMPCREG, BDFCMPSER, BDFCMPSSE, BDFCMPOPE, BDFCMPNAT"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYBDFCMP0_GetBuffer(rsSab, xYBDFCMP0)
        Nb = Nb + 1
        If Nb Mod 10 = 0 Then Call lstErr_ChangeLastItem(lstErr, cmdContext, "Export en cours : " & Nb & " enregistrements"): DoEvents

    wsExcel.Cells(Nb, 1) = xYBDFCMP0.BDFCMPSER
    wsExcel.Cells(Nb, 2) = xYBDFCMP0.BDFCMPSSE
    wsExcel.Cells(Nb, 3) = xYBDFCMP0.BDFCMPOPE
    wsExcel.Cells(Nb, 4) = xYBDFCMP0.BDFCMPNAT
    wsExcel.Cells(Nb, 5) = xYBDFCMP0.BDFCMPDOS
    wsExcel.Cells(Nb, 6) = xYBDFCMP0.BDFCMPMON
    wsExcel.Cells(Nb, 7) = xYBDFCMP0.BDFCMPDEV
    wsExcel.Cells(Nb, 8) = xYBDFCMP0.BDFCMPMONE
    wsExcel.Cells(Nb, 9) = xYBDFCMP0.BDFCMPDCRE
    wsExcel.Cells(Nb, 10) = xYBDFCMP0.BDFCMPDOPE
    wsExcel.Cells(Nb, 11) = xYBDFCMP0.BDFCMPCREG
    wsExcel.Cells(Nb, 12) = xYBDFCMP0.BDFCMPXDB
    wsExcel.Cells(Nb, 13) = xYBDFCMP0.BDFCMPXDBN
    wsExcel.Cells(Nb, 14) = xYBDFCMP0.BDFCMPXCR
    wsExcel.Cells(Nb, 15) = xYBDFCMP0.BDFCMPXCRN
    wsExcel.Cells(Nb, 16) = xYBDFCMP0.BDFCMPBBIC
    wsExcel.Cells(Nb, 17) = xYBDFCMP0.BDFCMPPAYS
    wsExcel.Cells(Nb, 18) = xYBDFCMP0.BDFCMPSTAT
    wsExcel.Cells(Nb, 19) = xYBDFCMP0.BDFCMPSTA
    wsExcel.Cells(Nb, 20) = xYBDFCMP0.BDFCMPUPDS
    wsExcel.Cells(Nb, 21) = xYBDFCMP0.BDFCMPUSR
    wsExcel.Cells(Nb, 22) = xYBDFCMP0.BDFCMPSEQ
    wsExcel.Cells(Nb, 23) = xYBDFCMP0.BDFCMP2008

        If mId$(xYBDFCMP0.BDFCMP2008, 1, 1) = "T" Or mId$(xYBDFCMP0.BDFCMP2008, 1, 1) = "X" Then
            iBackColor = 15
        Else
            Select Case mId$(xYBDFCMP0.BDFCMP2008, 2, 1)
                Case "A": iBackColor = 6
                Case "B": iBackColor = 4
                Case "C": iBackColor = 8
                Case "S": iBackColor = 4
                Case "T": iBackColor = 7
                Case Else: iBackColor = 3
                
            End Select
        End If
        wsExcel.Cells(Nb, 23).Interior.ColorIndex = iBackColor

    wsExcel.Cells(Nb, 24) = xYBDFCMP0.BDFCMPSABK
    wsExcel.Cells(Nb, 25) = xYBDFCMP0.BDFCMPMTK
    wsExcel.Cells(Nb, 26) = xYBDFCMP0.BDFCMPROUT
    wsExcel.Cells(Nb, 27) = xYBDFCMP0.BDFCMP50PI
    wsExcel.Cells(Nb, 28) = xYBDFCMP0.BDFCMP59PI


    rsSab.MoveNext
Loop
Set rsSab = Nothing


'____________________________________________________________________________________
wbExcel.SaveAs wFile

wbExcel.Close
appExcel.Quit

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
Call lstErr_AddItem(lstErr, cmdContext, "Export terminé : " & Nb & " enregistrements"): DoEvents

'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

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
        For I = fgSelect_arrIndex To 0 Step -1
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.LeftCol = 0
    End If
End If

End Sub
Private Sub fgSelect_Display()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
cmdPrint.Enabled = False
currentAction = "fgselect_Display"

For I = 1 To arrYBDFCMP0_Nb

        xYBDFCMP0 = arrYBDFCMP0(I)
    
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
Next I

Call lstErr_Clear(lstErr, cmdContext, "Nb enregistrements : " & fgSelect.Rows - 1): DoEvents
If fgSelect.Rows > 1 Then
'    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
    cmdPrint.Enabled = True
End If
fgSelect.Visible = True
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
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_1"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = True
fraSelect_Options_1.Enabled = True

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
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_2"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
txtSelect_AmjMin.Visible = True
txtSelect_AmjMin.Enabled = True
txtSelect_AmjMax.Visible = True
txtSelect_AmjMax.Enabled = True
Call DTPicker_Set(txtSelect_AmjMax, YBIATAB0_DATE_CAL_AP1)
Call DTPicker_Set(txtSelect_AmjMin, mId$(YBIATAB0_DATE_CAL_AP1, 1, 4) & "0101")

cmdSelect_Ok.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Private Sub lstSelect_Load_3()
Dim I As Long, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_3"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
txtSelect_AmjMin.Visible = True
txtSelect_AmjMin.Enabled = True
txtSelect_AmjMax.Visible = True
txtSelect_AmjMax.Enabled = True
Call DTPicker_Set(txtSelect_AmjMax, YBIATAB0_DATE_CAL_AP1)
Call DTPicker_Set(txtSelect_AmjMin, mId$(YBIATAB0_DATE_CAL_AP1, 1, 4) & "0101")

cmdSelect_Ok.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
Dim X As String, wColor As Long

On Error Resume Next
Select Case xYBDFCMP0.BDFCMPSTA
    Case Is = " ": wColor = vbBlue
    Case Else: wColor = vbGrayText
End Select

fgSelect.Col = 0: fgSelect.Text = xYBDFCMP0.BDFCMPSER & " " & xYBDFCMP0.BDFCMPSSE
fgSelect.CellForeColor = wColor
If xYBDFCMP0.BDFCMPOPE = "CDE" Or xYBDFCMP0.BDFCMPOPE = "CDI" Then
    X = xYBDFCMP0.BDFCMPOPE & " " & Format$(xYBDFCMP0.BDFCMPDOS, "#####0#") & " " & xYBDFCMP0.BDFCMPNAT
Else
    X = xYBDFCMP0.BDFCMPOPE & " " & xYBDFCMP0.BDFCMPNAT & " " & Format$(xYBDFCMP0.BDFCMPDOS, "#####0#")
End If
fgSelect.Col = 1: fgSelect.Text = X
fgSelect.CellForeColor = wColor
X = Format$(Abs(xYBDFCMP0.BDFCMPMON), "### ### ### ###.00")
fgSelect.Col = 2: fgSelect.Text = X
fgSelect.CellForeColor = vbRed
fgSelect.Col = 3: fgSelect.Text = xYBDFCMP0.BDFCMPDEV
fgSelect.CellForeColor = wColor
fgSelect.Col = 4: fgSelect.Text = xYBDFCMP0.BDFCMPCREG
fgSelect.CellForeColor = wColor
fgSelect.Col = 5: fgSelect.Text = xYBDFCMP0.BDFCMPBBIC
fgSelect.CellForeColor = wColor
fgSelect.Col = 6: fgSelect.Text = xYBDFCMP0.BDFCMP2008
fgSelect.CellForeColor = wColor

fgSelect.Col = 7: fgSelect.Text = dateIBM10(xYBDFCMP0.BDFCMPDCRE, True)
fgSelect.CellForeColor = wColor


fgSelect.Col = 8: fgSelect.Text = xYBDFCMP0.BDFCMPXDB & " " & xYBDFCMP0.BDFCMPXCR
fgSelect.CellForeColor = wColor

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
Dim wIndex As Integer
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    wIndex = Val(fgSelect.Text)
    Select Case lK
        Case 2: X = Format$(arrYBDFCMP0(wIndex).BDFCMPMON, "000000000000000.00")
        Case 3: X = arrYBDFCMP0(wIndex).BDFCMPDEV & Format$(arrYBDFCMP0(wIndex).BDFCMPMON, "000000000000000.00")
        Case 6: X = arrYBDFCMP0(wIndex).BDFCMPDCRE
    End Select
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub
Private Sub fgStatistiques_Display()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 1
fgStatistiques.Visible = False
fgStatistiques_Reset
fgStatistiques.Rows = 1
fgStatistiques.FormatString = fgStatistiques_FormatString
cmdPrint.Enabled = False
currentAction = "fgStatistiques_Display"


fgStatistiques.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Public Sub fgStatistiques_Reset()
fgStatistiques.Clear
fgStatistiques_Sort1 = 0: fgStatistiques_Sort2 = 0
fgStatistiques_Sort1_Old = -1
fgStatistiques_RowDisplay = 0: fgStatistiques_RowClick = 0
fgStatistiques_arrIndex = fgStatistiques.Cols - 1
blnfgStatistiques_DisplayLine = False
fgStatistiques_SortAD = 6
fgStatistiques.LeftCol = 0

End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
If fraUpdate.Visible Then fraUpdate.Visible = False: Exit Sub
If fgSelect.Visible Then fgSelect.Visible = False: cmdSelect_Ok.Caption = "Extraire les mouvements": Exit Sub
Unload Me
End Sub




Private Sub cboSelect_SQL_Click()
cmdSelect_SQL_K = mId$(cboSelect_SQL, 1, 1)
If blnControl Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    lstSelect.Visible = False
    txtSelect_AmjMin.Visible = False
    txtSelect_AmjMax.Visible = False
    fraSelect_Options_1.Visible = False
    fraUpdate.Visible = False
    fgStatistiques.Visible = False
    Select Case cmdSelect_SQL_K
        Case "1": lstSelect_Load_1
        Case "2", "4": lstSelect_Load_2
        Case "3": lstSelect_Load_3
        Case "5": lstSelect_Load_5
        Case "6": lstSelect_Load_6
        Case "7": lstSelect_Load_7
    End Select
    Me.Enabled = True: Me.MousePointer = 0
End If
End Sub


Private Sub chkSelect_BDFCMPDCRE_Click()
If chkSelect_BDFCMPDCRE = "1" Then
    If cmdSelect_SQL_K = "1" Then txtSelect_BDFCMPDCRE.Visible = True
    txtSelect_BDFCMPDCRE_Max.Visible = True
Else
    txtSelect_BDFCMPDCRE.Visible = False
    txtSelect_BDFCMPDCRE_Max.Visible = False
End If


End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdContext

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

blnAuto = False
blnAuto_Ok = False
lstSelect.Visible = False
cmdSelect_Ok.Caption = "Extraire les mouvements"

libRéférenceInterne = ""
cboSelect_SQL.ListIndex = 0
fgSelect.Visible = False
Call DTPicker_Set(txtSelect_AmjMin, YBIATAB0_DATE_CPT_JS1 - 10000)
txtSelect_AmjMax = txtSelect_AmjMin
txtSelect_BDFCMPDCRE = txtSelect_AmjMin
txtSelect_BDFCMPDCRE_Max = txtSelect_AmjMin
'Call DTPicker_Set(txtSelect_BDFCMPDCRE, YBIATAB0_DATE_CPT_JS1)
'Call DTPicker_Set(txtSelect_BDFCMPDCRE_Max, YBIATAB0_DATE_CPT_JS1)
fraUpdate.Visible = False
cboSelect_SQL.ListIndex = 1
blnControl = True
cboSelect_SQL.ListIndex = 0
End Sub
Public Sub Form_Init()
Dim xSql As String
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

SSTab1.Tab = 0

blnControl = False

fgSelect_FormatString = fgSelect.FormatString
fgStatistiques_FormatString = fgStatistiques.FormatString
cmdSelect_Ok.Visible = False
fraSelect_Options_1.Visible = False
'?fraBDFCMPSTAT.Visible = False
'?lstBDFCMPSTAT_Display.Enabled = False
'?lstBDFCMPSTAT_Display.ForeColor = vbMagenta
txtSelect_BDFCMPDCRE.Visible = False
txtSelect_BDFCMPDCRE_Max.Visible = False
cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1 - Consultation des opérations"
'2007 cboSelect_SQL.AddItem "2 - Déclaration cartographie 2007"
'2007 cboSelect_SQL.AddItem "3 - Comptage Opération/Devise"
cboSelect_SQL.AddItem "4 - Déclaration cartographie 2008"
' cboSelect_SQL.AddItem "5 - Importation TRF (ZCHGOPE0)"
'2007 cboSelect_SQL.AddItem "6 - Importation CDO (ZMOUVEA0)"
'2007 cboSelect_SQL.AddItem "7 - Importation RD% (ZENCREG0)"
cboSelect_SQL.AddItem "V - export comptage V2008"
cboSelect_SQL.AddItem "E - export YBDFCMP0"

Call cbo_Load("BDF", "BDFCMP", cboUpdate_BDFCMPSTAT, 3)
cboUpdate_BDFCMP2008.Clear
xSql = "select BDFCMP2008 from " & paramIBM_Library_SABSPE & ".YBDFCMP0 group by BDFCMP2008 " _
    & " order by BDFCMP2008"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboUpdate_BDFCMP2008.AddItem rsSab("BDFCMP2008")
    rsSab.MoveNext

Loop




X = "select * from ElpTable where SNN = 0" _
    & " and id = 'PAYS'" _
    & " and K1 = 'UE'"
arrPAYS_UE_Nb = 0
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    arrPAYS_UE_Nb = arrPAYS_UE_Nb + 1
   arrPAYS_UE(arrPAYS_UE_Nb) = Trim(rsMDB("K2"))
    rsMDB.MoveNext
Loop


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


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Dim Msg As String
Dim I As Integer

Me.Enabled = False: Me.MousePointer = vbHourglass
    Select Case cmdSelect_SQL_K
        Case "2": cmdPrint_Ok_2
        Case "3": cmdPrint_Ok_3
        Case "4": cmdPrint_Ok_4
    End Select

Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

Me.Enabled = False: Me.MousePointer = vbHourglass
    blnOk = Not fgSelect.Visible
Call lstErr_Clear(lstErr, cmdContext, "> SAB_CDR_cmdSelect_Ok ........"): DoEvents
cmdSelect_Ok.Visible = False
fraUpdate.Visible = False
fraSelect_Options_1.Enabled = False
txtSelect_AmjMin.Enabled = False
txtSelect_AmjMax.Enabled = False
fgSelect.Clear
DoEvents
If blnOk Then
    cmdSelect_Ok.Caption = "Modifier les options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    lstSelect.BackColor = &H8000000F
    Call usrColor_Container(lstSelect, lstSelect.BackColor)
    Select Case cmdSelect_SQL_K
        Case "1": cmdSelect_SQL
        Case "2": cmdSelect_SQL_2
        Case "3": cmdSelect_SQL_3
        Case "4": cmdSelect_SQL_4
        Case "5": cmdSelect_SQL_5
        Case "6": cmdSelect_SQL_6
        Case "7": cmdSelect_SQL_7
        Case "V": YBDFCMP0_V2008_Export
        Case "E": YBDFCMP0_Export
    End Select

    fgSelect.Enabled = True
Else
    cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
    cmdSelect_Ok.BackColor = &HC0FFC0
    lstSelect.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(lstSelect, lstSelect.BackColor)
    fgSelect.Visible = False
    fgSelect.Enabled = False
    fraSelect_Options_1.Enabled = True
    txtSelect_AmjMin.Enabled = True
    txtSelect_AmjMax.Enabled = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_CDR_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
cmdSelect_Ok.Visible = True

End Sub


Private Sub cmdSelect_SQL()
Dim V
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL"): DoEvents

currentAction = "cmdSelect_SQL"

Set rsSab = Nothing
Call DTPicker_Control(txtSelect_BDFCMPDCRE, wAmjMin)
Call DTPicker_Control(txtSelect_BDFCMPDCRE_Max, wAmjMax)

If chkSelect_BDFCMPDCRE = "1" Then
    xWhere = xWhere & " and BDFCMPDCRE >= " & wAmjMin _
                    & " and BDFCMPDCRE <= " & wAmjMax
End If
X = Trim(txtSelect_BDFCMPOPE)
If X <> "" Then xWhere = xWhere & " and BDFCMPOPE like '%" & X & "%'"
X = Trim(txtSelect_BDFCMPNAT)
If X <> "" Then xWhere = xWhere & " and BDFCMPNAT like '%" & X & "%'"
X = Trim(txtSelect_BDFCMPSTAT)
If X <> "" Then xWhere = xWhere & " and BDFCMP2008 like '%" & X & "%'"
If chkSelect_BDFCMPSTA = "0" Then xWhere = xWhere & " and BDFCMPSTA = ' '"

xWhere = Replace(xWhere, "and", "where", , 1)
arrYBDFCMP0_SQL xWhere & " order by BDFCMPOPE,BDFCMPNAT,BDFCMPDOS"
    
fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_2()
Dim V
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
Dim wNb As Long, wBDFCMPMONE As Currency, wBDFCMPSTAT As String, xBDFCMPSTAT As String

On Error GoTo Error_Handler

ReDim arrYBDFCMP0(100)
fgStatistiques_Display
fgStatistiques.Rows = 1
fgStatistiques.Row = 0
wNb = 0: wBDFCMPMONE = 0: wBDFCMPSTAT = ""
Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL"): DoEvents

currentAction = "cmdSelect_SQL"
Call DTPicker_Control(txtSelect_AmjMin, wAmjMin)
Call DTPicker_Control(txtSelect_AmjMax, wAmjMax)

xWhere = " where  BDFCMPDCRE >= " & wAmjMin _
                & " and BDFCMPDCRE <= " & wAmjMax _
       & " order by BDFCMPSTAT"
'xWhere = " where  BDFCMPSTA = ' ' and BDFCMPDCRE >= " & wAmjMin _
'                & " and BDFCMPDCRE <= " & wAmjMax _
'       & " order by BDFCMPSTAT"

'xWhere = " where  BDFCMPSTA = ' ' order by BDFCMPSTAT"
    
Set rsSab = Nothing

xSql = "select BDFCMPMONE,BDFCMPSTAT from " & paramIBM_Library_SABSPE & ".YBDFCMP0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    xBDFCMPSTAT = rsSab("BDFCMPSTAT")
    If wBDFCMPSTAT <> xBDFCMPSTAT Then
        If wNb > 0 Then Call cmdSelect_SQL_2_Display(wBDFCMPSTAT, wBDFCMPMONE, wNb)
        wBDFCMPSTAT = xBDFCMPSTAT
        wNb = 1
        wBDFCMPMONE = CCur(rsSab("BDFCMPMONE"))
    Else
        wNb = wNb + 1
        wBDFCMPMONE = wBDFCMPMONE + CCur(rsSab("BDFCMPMONE"))
    End If
    rsSab.MoveNext

Loop
If wNb > 0 Then Call cmdSelect_SQL_2_Display(wBDFCMPSTAT, wBDFCMPMONE, wNb)
cmdPrint.Enabled = True
    
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_4()
Dim V
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
Dim wNb As Long, wBDFCMPMONE As Currency, wBDFCMP2008 As String, xBDFCMP2008 As String

On Error GoTo Error_Handler

ReDim arrYBDFCMP0(100)
fgStatistiques_Display
fgStatistiques.Rows = 1
fgStatistiques.Row = 0
wNb = 0: wBDFCMPMONE = 0: wBDFCMP2008 = ""
Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL"): DoEvents

currentAction = "cmdSelect_SQL"
Call DTPicker_Control(txtSelect_AmjMin, wAmjMin)
Call DTPicker_Control(txtSelect_AmjMax, wAmjMax)

xWhere = " where  BDFCMPDCRE >= " & wAmjMin _
                & " and BDFCMPDCRE <= " & wAmjMax _
       & " order by BDFCMP2008"
'xWhere = " where  BDFCMPSTA = ' ' and BDFCMPDCRE >= " & wAmjMin _
'                & " and BDFCMPDCRE <= " & wAmjMax _
'       & " order by BDFCMP2008"

'xWhere = " where  BDFCMPSTA = ' ' order by BDFCMP2008"
    
Set rsSab = Nothing

xSql = "select BDFCMPMONE,BDFCMP2008 from " & paramIBM_Library_SABSPE & ".YBDFCMP0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    xBDFCMP2008 = rsSab("BDFCMP2008")
    If wBDFCMP2008 <> xBDFCMP2008 Then
        If wNb > 0 Then
                Call cmdSelect_SQL_4_Display(wBDFCMP2008, wBDFCMPMONE, wNb)
        End If
        If mId$(wBDFCMP2008, 1, 3) <> mId$(xBDFCMP2008, 1, 3) Then
            fgStatistiques.Rows = fgStatistiques.Rows + 1
            fgStatistiques.Row = fgStatistiques.Rows - 1

            fgStatistiques.Col = 0
            fgStatistiques.Text = mId$(xBDFCMP2008, 1, 3) & "-1-2-3"
        End If
        wBDFCMP2008 = xBDFCMP2008
        wNb = 1
        wBDFCMPMONE = CCur(rsSab("BDFCMPMONE"))
    Else
        wNb = wNb + 1
        wBDFCMPMONE = wBDFCMPMONE + CCur(rsSab("BDFCMPMONE"))
    End If
    rsSab.MoveNext

Loop
If wNb > 0 Then Call cmdSelect_SQL_4_Display(wBDFCMP2008, wBDFCMPMONE, wNb)
cmdPrint.Enabled = True
    
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub cmdSelect_SQL_3()
Dim V
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
Dim wNb As Long, wBDFCMPMON As Currency
Dim wBDFCMPOPE As String, wBDFCMPDEV As String, wBDFCMPSTAT As String
Dim xBDFCMPOPE As String, xBDFCMPDEV As String, xBDFCMPSTAT As String
Dim xDisplay As String
On Error GoTo Error_Handler
ReDim arrYBDFCMP0(500)
fgStatistiques_Display
fgStatistiques.Rows = 1
fgStatistiques.Row = 0
wNb = 0: wBDFCMPMON = 0: wBDFCMPSTAT = ""
Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL"): DoEvents

currentAction = "cmdSelect_SQL"
Call DTPicker_Control(txtSelect_AmjMin, wAmjMin)
Call DTPicker_Control(txtSelect_AmjMax, wAmjMax)

xWhere = " where  BDFCMPDCRE >= " & wAmjMin _
                & " and BDFCMPDCRE <= " & wAmjMax _
       & " order by BDFCMPOPE,BDFCMPDEV,BDFCMPSTAT"
    
Set rsSab = Nothing

xSql = "select BDFCMPOPE,BDFCMPDEV,BDFCMPMON,BDFCMPSTAT from " & paramIBM_Library_SABSPE & ".YBDFCMP0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    xBDFCMPOPE = rsSab("BDFCMPOPE")
    xBDFCMPDEV = rsSab("BDFCMPDEV")
    xBDFCMPSTAT = rsSab("BDFCMPSTAT")
    If wBDFCMPOPE <> xBDFCMPOPE _
    Or wBDFCMPDEV <> xBDFCMPDEV _
    Or wBDFCMPSTAT <> xBDFCMPSTAT Then
        If wNb > 0 Then Call cmdSelect_SQL_3_Display(wBDFCMPOPE, wBDFCMPDEV, wBDFCMPSTAT, wBDFCMPMON, wNb)
        wBDFCMPOPE = xBDFCMPOPE
        wBDFCMPDEV = xBDFCMPDEV
        wBDFCMPSTAT = xBDFCMPSTAT
        xDisplay = wBDFCMPOPE & " " & wBDFCMPDEV & " " & wBDFCMPSTAT
        wNb = 1
        wBDFCMPMON = CCur(rsSab("BDFCMPMON"))
    Else
        wNb = wNb + 1
        wBDFCMPMON = wBDFCMPMON + CCur(rsSab("BDFCMPMON"))
    End If
    rsSab.MoveNext

Loop
If wNb > 0 Then Call cmdSelect_SQL_3_Display(wBDFCMPOPE, wBDFCMPDEV, wBDFCMPSTAT, wBDFCMPMON, wNb)
cmdPrint.Enabled = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_5()
Dim V
On Error GoTo Error_Handler

MsgBox "CMP 2008 : remplacé par les pgm cobol YBDFCMP0_B et YBDFCMP-S (argument '2008')", vbCritical, "BDFCMP : ZCHGOPE0"
Exit Sub


Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_5"): DoEvents

currentAction = "cmdSelect_SQL_5"
Call DTPicker_Control(txtSelect_AmjMin, wAmjMin)
Call DTPicker_Control(txtSelect_AmjMax, wAmjMax)
If wAmjMin = "00000000" Then
    MsgBox "Préciser la date", vbInformation, "Import des MAD à une date"
    Exit Sub
End If
    
cmdSelect_SQL_5_ZCHGOPE0

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_6()
Dim V
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_6"): DoEvents

currentAction = "cmdSelect_SQL_6"
Call DTPicker_Control(txtSelect_AmjMin, wAmjMin)
Call DTPicker_Control(txtSelect_AmjMax, wAmjMax)
If wAmjMin = "00000000" Then
    MsgBox "Préciser la date", vbInformation, "Import des TRF à une date"
    Exit Sub
End If
    
cmdSelect_SQL_6_ZMOUVEA0

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub cmdSelect_SQL_7()
Dim V
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_7"): DoEvents

currentAction = "cmdSelect_SQL_7"
Call DTPicker_Control(txtSelect_AmjMin, wAmjMin)
Call DTPicker_Control(txtSelect_AmjMax, wAmjMax)
If wAmjMin = "00000000" Then
    MsgBox "Préciser la date", vbInformation, "Import des TRF à une date"
    Exit Sub
End If
    
cmdSelect_SQL_7_ZENCREG0

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_5_ZCHGOPE0()
Dim V
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String, X As String
Dim blnOk As Boolean
On Error GoTo Error_Handler

    
xWhere = "where BDFCMPDCRE >= " & wAmjMin - 19000000 _
        & " and  BDFCMPDCRE <= " & wAmjMax - 19000000 _
        & " order by BDFCMPOPE,BDFCMPNAT,BDFCMPDOS"
arrYBDFCMP0_SQL xWhere

Set rsSab = Nothing
xWhere = " where CHGOPECRE >= " & wAmjMin - 19000000 & " and  CHGOPECRE <= " & wAmjMax - 19000000 _
       & " and CHGOPEANN = ' '"

xSql = "select * from " & paramIBM_Library_SAB & ".ZCHGOPE0 " & xWhere & " order by CHGOPECRE,CHGOPEOPE,CHGOPENAT,CHGOPEDOS"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    blnOk = False
        blnOk = True
        xYBDFCMP0.BDFCMPSER = rsSab("CHGOPESER")
        xYBDFCMP0.BDFCMPSSE = rsSab("CHGOPESSE")
        xYBDFCMP0.BDFCMPDOS = rsSab("CHGOPEDOS")
        xYBDFCMP0.BDFCMPOPE = rsSab("CHGOPEOPE")
        xYBDFCMP0.BDFCMPNAT = rsSab("CHGOPENAT")
        For K = 1 To arrYBDFCMP0_Nb
            If xYBDFCMP0.BDFCMPDOS < arrYBDFCMP0(K).BDFCMPDOS Then Exit For
            If xYBDFCMP0.BDFCMPDOS = arrYBDFCMP0(K).BDFCMPDOS Then
                If xYBDFCMP0.BDFCMPOPE = arrYBDFCMP0(K).BDFCMPOPE _
                And xYBDFCMP0.BDFCMPNAT = arrYBDFCMP0(K).BDFCMPNAT _
                And xYBDFCMP0.BDFCMPSER = arrYBDFCMP0(K).BDFCMPSER _
                And xYBDFCMP0.BDFCMPSSE = arrYBDFCMP0(K).BDFCMPSSE Then blnOk = False: Exit For
            End If
        Next K
        If blnOk Then YBDFCMP0_Add_ZCHGOPE0
    
    rsSab.MoveNext

Loop
    
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub cmdSelect_SQL_5_ANN()
Dim V
Dim xSql As String, K As Long, Nb As Long
Dim xWhere As String, xAnd As String, X As String
Dim blnOk As Boolean
On Error GoTo Error_Handler

K = 0
Set rsSab = Nothing
xWhere = " where CHGOPECRE >= 1050000 and  CHGOPECRE <= 1059999" _
       & " and CHGOPEANN = 'A'"

xSql = "select CHGOPEOPE,CHGOPENAT,CHGOPEDOS from " & paramIBM_Library_SAB & ".ZCHGOPE0 " & xWhere
Set rsSab = cnsab.Execute(xSql)
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

Do While Not rsSab.EOF
    K = K + 1
    xSql = "delete from " & paramIBM_Library_SABSPE & ".YBDFCMP0" _
    & " where BDFCMPOPE = '" & rsSab("CHGOPEOPE") & "'" _
    & " and BDFCMPNAT = '" & rsSab("CHGOPENAT") & "'" _
    & " and BDFCMPDOS = " & rsSab("CHGOPEDOS")
    Debug.Print xSql
    Call FEU_ROUGE
    Set rsSabX = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
   If Nb = 0 Then
        MsgBox xSql, vbCritical, "Inconnu"
    End If
 

    rsSab.MoveNext

Loop
V = cnSAB_Transaction("Commit")
MsgBox "Nb " & K, vbInformation, "ANN"
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_7_ZENCREG0()
Dim V
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String, X As String
Dim blnOk As Boolean, blnMoveNext As Boolean
On Error GoTo Error_Handler

    

Set rsSab = Nothing
xWhere = " where ENCREGCOP like 'RD%' and ENCREGDCR >= " & wAmjMin - 19000000 & " and  ENCREGDCR <= " & wAmjMax - 19000000

xSql = "select * from " & paramIBM_Library_SAB & ".ZENCREG0 " & xWhere & " order by ENCREGCOP,ENCREGDOS,ENCREGREG,ENCREGSEN"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    blnOk = True
    blnMoveNext = True
    mENCREGCL5_D = "": mENCREGBC5_D = "": mENCREGLC5_D = ""
    mENCREGCL5_C = "": mENCREGBC5_C = "": mENCREGLC5_C = ""

    rsYBDFCMP0_Init xYBDFCMP0
    If rsSab("ENCREGDAN") <> 0 Then blnOk = False
    xYBDFCMP0.BDFCMPSER = rsSab("ENCREGSER")
    xYBDFCMP0.BDFCMPSSE = rsSab("ENCREGSSE")
    xYBDFCMP0.BDFCMPOPE = rsSab("ENCREGCOP")
    xYBDFCMP0.BDFCMPNAT = rsSab("ENCREGREG")
    xYBDFCMP0.BDFCMPDOS = rsSab("ENCREGDOS")
    xYBDFCMP0.BDFCMPMON = rsSab("ENCREGMOD")
    xYBDFCMP0.BDFCMPDEV = rsSab("ENCREGDEV")
    xYBDFCMP0.BDFCMPMONE = rsSab("ENCREGMOB")
    xYBDFCMP0.BDFCMPDCRE = rsSab("ENCREGDCR") + 19000000
    xYBDFCMP0.BDFCMPDOPE = rsSab("ENCREGDEN") + 19000000
    xYBDFCMP0.BDFCMPCREG = rsSab("ENCREGMOR")
    If rsSab("ENCREGSEN") = "D" Then
        xYBDFCMP0.BDFCMPXDB = rsSab("ENCREGCOM")
        mENCREGCL5_D = rsSab("ENCREGCL5")
        mENCREGBC5_D = rsSab("ENCREGBC5")
        mENCREGLC5_D = rsSab("ENCREGLC5")

    Else
        xYBDFCMP0.BDFCMPXCR = rsSab("ENCREGCOM")
        mENCREGCL5_C = rsSab("ENCREGCL5")
        mENCREGBC5_C = rsSab("ENCREGBC5")
        mENCREGLC5_C = rsSab("ENCREGLC5")
   End If
'______________________________________________________
    

    rsSab.MoveNext
    
    If xYBDFCMP0.BDFCMPDOS <> rsSab("ENCREGDOS") Then
        blnOk = False: blnMoveNext = False
        MsgBox "manque suite DOS " & xYBDFCMP0.BDFCMPDOS
    End If
    X = rsSab("ENCREGREG")
    If Trim(xYBDFCMP0.BDFCMPNAT) <> X Then
        blnOk = False: blnMoveNext = False
        MsgBox "manque suite DOS/REG  " & xYBDFCMP0.BDFCMPDOS
    End If
    If rsSab("ENCREGDAN") <> 0 Then blnOk = False
    X = rsSab("ENCREGMOR")
    If X = "SWF" Or X = "TGT" Or X = "SNP" Then xYBDFCMP0.BDFCMPCREG = X

    If rsSab("ENCREGSEN") = "D" Then
        xYBDFCMP0.BDFCMPXDB = rsSab("ENCREGCOM")
        mENCREGCL5_D = rsSab("ENCREGCL5")
        mENCREGBC5_D = rsSab("ENCREGBC5")
        mENCREGLC5_D = rsSab("ENCREGLC5")
    Else
        xYBDFCMP0.BDFCMPXCR = rsSab("ENCREGCOM")
        mENCREGCL5_C = rsSab("ENCREGCL5")
        mENCREGBC5_C = rsSab("ENCREGBC5")
        mENCREGLC5_C = rsSab("ENCREGLC5")
    End If

    If blnOk Then YBDFCMP0_Add_ZENCREG0
    
    If blnMoveNext Then rsSab.MoveNext

Loop
    
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_6_ZMOUVEA0()
Dim V
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String, X As String
Dim blnOk As Boolean

Dim mMOUVEMMON As Currency
On Error GoTo Error_Handler

    
xWhere = "where BDFCMPDCRE >= " & wAmjMin - 19000000 _
        & " and  BDFCMPDCRE <= " & wAmjMax - 19000000 _
        & " order by BDFCMPOPE,BDFCMPNAT,BDFCMPDOS"
arrYBDFCMP0_SQL xWhere

Set rsSab = Nothing
Set rsSabX = Nothing
xWhere = " where MOUVEMETA = 1 and MOUVEMPLA = 1" _
    & " and MOUVEMCOM like '388980%' and MOUVEMEVE = 'RGL' and MOUVEMDTR >= " & wAmjMin - 19000000 & " and  MOUVEMDTR <= " & wAmjMax - 19000000

xSql = "select MOUVEMPIE,MOUVEMMON,MOUVEMNUM from " & paramIBM_Library_SAB & ".ZMOUVEMA " & xWhere & " order by MOUVEMPIE"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    mMOUVEMMON = Abs(rsSab("MOUVEMMON"))                  '!!! montant Débit pour SQL ZCDOREG0 (plusieurs paiements)
'If rsSab("MOUVEMNUM") = 79131 Then
'    Debug.Print "Debug"
'End If

    xSql = "select * from " & paramIBM_Library_SAB & ".ZMOUVEMG" _
    & " where MOUVEMETA = 1 and MOUVEMPLA = 1" _
    & " and MOUVEMPIE = " & rsSab("MOUVEMPIE") & " and MOUVEMMON < 0 order by MOUVEMMON"
    Set rsSabX = cnsab.Execute(xSql)

    If Not rsSabX.EOF Then
            blnOk = True
            rsYBDFCMP0_Init xYBDFCMP0

            xYBDFCMP0.BDFCMPSER = rsSabX("MOUVEMSER")
            xYBDFCMP0.BDFCMPSSE = rsSabX("MOUVEMSSE")
            xYBDFCMP0.BDFCMPDOS = rsSabX("MOUVEMNUM")
            xYBDFCMP0.BDFCMPOPE = rsSabX("MOUVEMOPE")
            xYBDFCMP0.BDFCMPMON = Abs(rsSabX("MOUVEMMON"))
            xYBDFCMP0.BDFCMPDOPE = rsSabX("MOUVEMDCO") + 19000000
            xYBDFCMP0.BDFCMPDCRE = rsSabX("MOUVEMDTR") + 19000000
            xYBDFCMP0.BDFCMPXCR = rsSabX("MOUVEMCOM")
'If xYBDFCMP0.BDFCMPDOS = 79131 Then
'    Debug.Print xYBDFCMP0.BDFCMPDOS
'End If

            xSql = "select * from " & paramIBM_Library_SAB & ".ZLIBEL0" _
            & " where LIBELETA = 1 " _
            & " and LIBELPIE = " & rsSabX("MOUVEMPIE") _
            & " and LIBELECR = " & rsSabX("MOUVEMECR") & " and LIBELNUM = 1"
            Set rsSabX = cnsab.Execute(xSql)
            If Not rsSabX.EOF Then
                xYBDFCMP0.BDFCMPNAT = Right$(Trim(rsSabX("LIBELLIB")), 2)
           Else
                xYBDFCMP0.BDFCMPNAT = "XX"
            End If
            
        For K = 1 To arrYBDFCMP0_Nb
                If xYBDFCMP0.BDFCMPDOS < arrYBDFCMP0(K).BDFCMPDOS Then Exit For
                If xYBDFCMP0.BDFCMPDOS = arrYBDFCMP0(K).BDFCMPDOS Then
                    If xYBDFCMP0.BDFCMPOPE = arrYBDFCMP0(K).BDFCMPOPE _
                    And xYBDFCMP0.BDFCMPNAT = arrYBDFCMP0(K).BDFCMPNAT _
                    And xYBDFCMP0.BDFCMPSER = arrYBDFCMP0(K).BDFCMPSER _
                    And xYBDFCMP0.BDFCMPSSE = arrYBDFCMP0(K).BDFCMPSSE Then blnOk = False: Exit For
                End If
            Next K
           If blnOk Then YBDFCMP0_Add_ZMOUVEA0 mMOUVEMMON
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




Private Sub arrYBDFCMP0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrYBDFCMP0(501)
arrYBDFCMP0_Max = 500: arrYBDFCMP0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBDFCMP0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYBDFCMP0_GetBuffer(rsSab, xYBDFCMP0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgselect_Display"
        '' Exit Sub
     Else
         arrYBDFCMP0_Nb = arrYBDFCMP0_Nb + 1
         If arrYBDFCMP0_Nb > arrYBDFCMP0_Max Then
             arrYBDFCMP0_Max = arrYBDFCMP0_Max + 50
             ReDim Preserve arrYBDFCMP0(arrYBDFCMP0_Max)
         End If
         
         arrYBDFCMP0(arrYBDFCMP0_Nb) = xYBDFCMP0
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

Private Sub cmdUpdate_Annuler_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents

newYBDFCMP0 = oldYBDFCMP0
If newYBDFCMP0.BDFCMPSTA = " " Then
    newYBDFCMP0.BDFCMPSTA = "A"
    newYBDFCMP0.BDFCMPSTAT = "999"
Else
    newYBDFCMP0.BDFCMPSTA = " "
End If


    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents
    V = cmdUpdate_Ok_Transaction
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        arrYBDFCMP0(arrYBDFCMP0_Index) = newYBDFCMP0
        xYBDFCMP0 = newYBDFCMP0
        fgSelect_DisplayLine arrYBDFCMP0_Index
        fraUpdate.Visible = False

    Else
        MsgBox V, vbCritical, Me.Name & " : cmdUpdate_Ok"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdUpdate_Ok_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents

If IsNull(fraUpdate_Control) Then
    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents
    V = cmdUpdate_Ok_Transaction
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        arrYBDFCMP0(arrYBDFCMP0_Index) = newYBDFCMP0
        xYBDFCMP0 = newYBDFCMP0
        fgSelect_DisplayLine arrYBDFCMP0_Index
        fraUpdate.Visible = False
    Else
        MsgBox V, vbCritical, Me.Name & " : cmdUpdate_Ok"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdUpdate_Quit_Click()
fraUpdate.Visible = False

End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
Me.Enabled = False
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
        Select Case fgSelect.Col
            Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_SortX 0
            Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
            Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_SortX 2
            Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_SortX 3
            Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
            Case 5: fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_Sort
            Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
            'Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
            Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
           Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
        End Select
Else
    If fgSelect.Rows > 1 Then
        fgSelect.Col = fgSelect_arrIndex:  arrYBDFCMP0_Index = CLng(fgSelect.Text)
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        xYBDFCMP0 = arrYBDFCMP0(arrYBDFCMP0_Index)
        oldYBDFCMP0 = xYBDFCMP0
        fraUpdate_Display
   End If
End If
Me.Enabled = True
End Sub


Public Function fraUpdate_Control()
Dim blnUpdate_Control As Boolean
Dim X As String
blnUpdate_Control = True
Call lstErr_AddItem(lstErr, cmdContext, ">_________Contrôle des données "): DoEvents
newYBDFCMP0 = oldYBDFCMP0


X = Trim(cboUpdate_BDFCMPSTAT)
If X = "" Then
    blnUpdate_Control = False
    cboUpdate_BDFCMPSTAT.BackColor = errUsr.BackColor
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le code Grille")
Else
    cboUpdate_BDFCMPSTAT.BackColor = txtUsr.BackColor
End If
newYBDFCMP0.BDFCMP2008 = mId$(X, 1, 3)
X = Trim(cboUpdate_BDFCMP2008)
If X = "" Then
    blnUpdate_Control = False
    cboUpdate_BDFCMP2008.BackColor = errUsr.BackColor
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le code Grille")
Else
    cboUpdate_BDFCMP2008.BackColor = txtUsr.BackColor
End If
newYBDFCMP0.BDFCMP2008 = mId$(X, 1, 4)

If blnUpdate_Control Then
    fraUpdate_Control = Null
Else
    fraUpdate_Control = "?_________fraUpdate_Control"
End If
End Function

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

Call BiaPgmAut_Init(mId$(Msg, 1, 12), BIA_BDFCMP_Aut)

blnSetfocus = True
Form_Init
blnAuto = False

Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case Else: blnAuto = False
End Select


End Sub


Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
'    If fraUpdate.Visible _
'   And fraUpdate_B.Enabled _
'    And cmdUpdate_Ok.Enabled Then cmdUpdate_Ok_Click: Exit Sub
Else
    If currentAction = "" Then
        If SSTab1.Tab > 0 Then
            SSTab1.Tab = 0
        Else
           'SendKeys "{TAB}"
           ' cmdSelect_Click
        End If
    End If
End If
End Sub









Private Sub mnuPrint0_All_Click()
Dim I As Long, K As Long
Me.Enabled = False: Me.MousePointer = vbHourglass
    
For I = 1 To arrYBDFCMP0_Nb
    fgSelect.Row = I
    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
    xYBDFCMP0 = arrYBDFCMP0(K)
    'prtSAB_CDR_Monitor xYBDFCMP0
Next I

Me.Show

Me.Enabled = True: Me.MousePointer = 0



End Sub




Public Sub cmdPrint_Ok()
Dim K As Long, X As String, xSql As String
Dim wMOUVEMCOM As String
lstSelect.Visible = False


End Sub
Public Sub cmdPrint_Ok_1()
Dim K As Long, X As String
Dim wIndex As Integer

fgSelect.Visible = False

fgSelect.Visible = True


End Sub


Public Function YBDFCMP0_Add_ZCHGOPE0()
Dim V, X As String, xSql As String, xWhere As String
Dim Nb As Long
Dim mMsgBox As String
Dim wDev As String, curDev As Currency, curEur As Currency, wAmj As Long
Dim wBDFCMPNAT As String
Dim blnEmis As Boolean, mBDFCMPSTAT_Ligne As String
Dim wCours As Double

On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "YBDFCMP0_Add_ZCHGOPE0"
'-------------------------------------------------------

YBDFCMP0_Add_ZCHGOPE0 = Null
mMsgBox = xYBDFCMP0.BDFCMPOPE & " " & xYBDFCMP0.BDFCMPNAT & " " & xYBDFCMP0.BDFCMPDOS
'wCours_AMJ = 0: wCours = 0: wCours_Dev = ""
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

Call lstErr_AddItem(lstErr, cmdContext, "Ajout MAD : " & mMsgBox): DoEvents
'________________________________________________________________________________
rsYBDFCMP0_Init meYBDFCMP0
meYBDFCMP0.BDFCMPSER = rsSab("CHGOPESER")
meYBDFCMP0.BDFCMPSSE = rsSab("CHGOPESSE")
meYBDFCMP0.BDFCMPOPE = rsSab("CHGOPEOPE")
meYBDFCMP0.BDFCMPDOS = rsSab("CHGOPEDOS")
meYBDFCMP0.BDFCMPNAT = rsSab("CHGOPENAT")
meYBDFCMP0.BDFCMPDCRE = CLng(rsSab("CHGOPECRE")) + 19000000
wAmj = rsSab("CHGOPEENG")
meYBDFCMP0.BDFCMPDOPE = wAmj + 19000000
'meYBDFCMP0.BDFCMPCREG = "???"
'meYBDFCMP0.BDFCMPXDB = "???"
'meYBDFCMP0.BDFCMPXDBN = "???"
'meYBDFCMP0.BDFCMPXCR = "???"
'meYBDFCMP0.BDFCMPXCRN = "???"
'meYBDFCMP0.BDFCMPBBIC = "???"
'meYBDFCMP0.BDFCMPPAYS = "??"
'meYBDFCMP0.BDFCMPSTAT = 0
'meYBDFCMP0.BDFCMPSTA = " "

wDev = rsSab("CHGOPEDE1")
If wDev = "EUR" Then
    curEur = rsSab("CHGOPEMO1")
    curDev = rsSab("CHGOPEMO2")
    If curDev = 0 Then
        curDev = curEur
    Else
        wDev = rsSab("CHGOPEDE2")
    End If
Else
    curEur = rsSab("CHGOPEMO2")
    curDev = rsSab("CHGOPEMO1")

End If

'________________________________________________________________________________
If curEur = 0 And wDev <> "EUR" Then
    wCours = rsZBASTAB0_Cours37(wDev, wAmj)
    If wCours > 0 Then curEur = curDev / wCours
End If
'________________________________________________________________________________

meYBDFCMP0.BDFCMPDEV = wDev
meYBDFCMP0.BDFCMPMON = curDev
meYBDFCMP0.BDFCMPMONE = curEur

Set rsSabX = Nothing
xWhere = " where CHGDETDOS = " & meYBDFCMP0.BDFCMPDOS _
       & " and CHGDETOPE = '" & meYBDFCMP0.BDFCMPOPE & "'" _
       & " and CHGDETTYP = 'P'" _
       & " and CHGDETSER = '" & rsSab("CHGOPESER") & "' and CHGDETSSE = '" & rsSab("CHGOPESSE") & "'"
            
xSql = "select CHGDETCP1,CHGDETSEN from " & paramIBM_Library_SAB & ".ZCHGDET0 " & xWhere
Set rsSabX = cnsab.Execute(xSql)


Do While Not rsSabX.EOF
    X = rsSabX("CHGDETSEN")
    If X = "C" Then
        meYBDFCMP0.BDFCMPXCR = rsSabX("CHGDETCP1")
    Else
        meYBDFCMP0.BDFCMPXDB = rsSabX("CHGDETCP1")
    End If
    
    rsSabX.MoveNext

Loop
'------------------------------------------

Set rsSabX = Nothing
xWhere = " where CHGMESDOS = " & meYBDFCMP0.BDFCMPDOS _
       & " and CHGMESOPE = '" & meYBDFCMP0.BDFCMPOPE & "'" _
       & " and CHGMESSEQ = '  '" _
       & " and CHGMESSER = '" & rsSab("CHGOPESER") & "' and CHGMESSSE = '" & rsSab("CHGOPESSE") & "'"
            
xSql = "select CHGMESBI3,CHGMESVIR from " & paramIBM_Library_SAB & ".ZCHGMES0 " & xWhere
Set rsSabX = cnsab.Execute(xSql)


If Not rsSabX.EOF Then
    meYBDFCMP0.BDFCMPBBIC = rsSabX("CHGMESBI3")
    X = rsSabX("CHGMESVIR")
    'If X <> "CCO" And X <> "INT" And X <> "   " Then
    meYBDFCMP0.BDFCMPCREG = X
    meYBDFCMP0.BDFCMPPAYS = mId$(meYBDFCMP0.BDFCMPBBIC, 5, 2)
End If


'------------------------------------------

Set rsSabX = Nothing
xWhere = " where COMPTECOM = '" & meYBDFCMP0.BDFCMPXCR & "'"
            
xSql = "select PLANCOPRO from " & paramIBM_Library_SABSPE & ".YBIACPT0 " & xWhere
Set rsSabX = cnsab.Execute(xSql)
If Not rsSabX.EOF Then meYBDFCMP0.BDFCMPXCRN = rsSabX("PLANCOPRO")

xWhere = " where COMPTECOM = '" & meYBDFCMP0.BDFCMPXDB & "'"
            
xSql = "select PLANCOPRO from " & paramIBM_Library_SABSPE & ".YBIACPT0 " & xWhere
Set rsSabX = cnsab.Execute(xSql)
If Not rsSabX.EOF Then meYBDFCMP0.BDFCMPXDBN = rsSabX("PLANCOPRO")

'-------------------------------------------------------------------------------------------------
If meYBDFCMP0.BDFCMPSER = "TC" Then meYBDFCMP0.BDFCMPSTAT = "911": meYBDFCMP0.BDFCMPSTA = "I": GoTo YBDFCMP0_Insert
If meYBDFCMP0.BDFCMPNAT = "CRE" Then meYBDFCMP0.BDFCMPSTAT = "911": meYBDFCMP0.BDFCMPSTA = "I": GoTo YBDFCMP0_Insert
If meYBDFCMP0.BDFCMPOPE = "CPT" Then
    If mId$(meYBDFCMP0.BDFCMPNAT, 1, 1) = "1" Then meYBDFCMP0.BDFCMPSTAT = "911": meYBDFCMP0.BDFCMPSTA = "I": GoTo YBDFCMP0_Insert
    If mId$(meYBDFCMP0.BDFCMPNAT, 1, 1) = "2" Then meYBDFCMP0.BDFCMPSTAT = "911": meYBDFCMP0.BDFCMPSTA = "I": GoTo YBDFCMP0_Insert
    If meYBDFCMP0.BDFCMPNAT = "RCL" Then meYBDFCMP0.BDFCMPSTAT = "911": meYBDFCMP0.BDFCMPSTA = "I": GoTo YBDFCMP0_Insert
End If

If meYBDFCMP0.BDFCMPCREG = "SNP" Or meYBDFCMP0.BDFCMPCREG = "TBF" Or meYBDFCMP0.BDFCMPCREG = "CAI" Or meYBDFCMP0.BDFCMPCREG = "INT" Then
     meYBDFCMP0.BDFCMPSTA = "I": meYBDFCMP0.BDFCMPSTAT = "921": GoTo YBDFCMP0_Insert
End If

If meYBDFCMP0.BDFCMPCREG = "CCO" Then meYBDFCMP0.BDFCMPSTAT = "211": GoTo YBDFCMP0_Insert
If mId$(meYBDFCMP0.BDFCMPXDBN, 1, 1) = "L" Then
    If meYBDFCMP0.BDFCMPXCRN = "INT" Then meYBDFCMP0.BDFCMPSTAT = "921": meYBDFCMP0.BDFCMPSTA = "I": GoTo YBDFCMP0_Insert
    If meYBDFCMP0.BDFCMPXCRN = "CAV" Or meYBDFCMP0.BDFCMPXCRN = "ASD" Then meYBDFCMP0.BDFCMPSTAT = "211": GoTo YBDFCMP0_Insert
    If mId$(meYBDFCMP0.BDFCMPXCRN, 1, 1) = "L" Then meYBDFCMP0.BDFCMPSTAT = "211": GoTo YBDFCMP0_Insert
End If
If mId$(meYBDFCMP0.BDFCMPXCRN, 1, 1) = "L" Then
    If meYBDFCMP0.BDFCMPXDBN = "INT" Then meYBDFCMP0.BDFCMPSTAT = "921": meYBDFCMP0.BDFCMPSTA = "I": GoTo YBDFCMP0_Insert
    If meYBDFCMP0.BDFCMPXDBN = "CAV" Or meYBDFCMP0.BDFCMPXDBN = "ASD" Then meYBDFCMP0.BDFCMPSTAT = "211": GoTo YBDFCMP0_Insert
    If mId$(meYBDFCMP0.BDFCMPXDBN, 1, 1) = "L" Then meYBDFCMP0.BDFCMPSTAT = "211": GoTo YBDFCMP0_Insert
End If

If meYBDFCMP0.BDFCMPXDBN = "NOS" Then
    blnEmis = True: mBDFCMPSTAT_Ligne = 25
Else
    If meYBDFCMP0.BDFCMPXCRN = "NOS" Then
        blnEmis = False: mBDFCMPSTAT_Ligne = 24
    Else
        meYBDFCMP0.BDFCMPSTAT = "931": GoTo YBDFCMP0_Insert
    End If
End If

If meYBDFCMP0.BDFCMPDEV <> "EUR" Then
    If meYBDFCMP0.BDFCMPCREG = "SWF" Or meYBDFCMP0.BDFCMPCREG = "TGT" Then
        meYBDFCMP0.BDFCMPSTAT = mBDFCMPSTAT_Ligne & "3": GoTo YBDFCMP0_Insert
    Else
        meYBDFCMP0.BDFCMPSTAT = "951": GoTo YBDFCMP0_Insert
    End If
End If

If meYBDFCMP0.BDFCMPPAYS = "  " Then meYBDFCMP0.BDFCMPSTAT = "941": GoTo YBDFCMP0_Insert
If meYBDFCMP0.BDFCMPPAYS = "FR" Then meYBDFCMP0.BDFCMPSTAT = mBDFCMPSTAT_Ligne & "1": GoTo YBDFCMP0_Insert
For I = 1 To arrPAYS_UE_Nb
    If meYBDFCMP0.BDFCMPPAYS = arrPAYS_UE(I) Then meYBDFCMP0.BDFCMPSTAT = mBDFCMPSTAT_Ligne & "2": GoTo YBDFCMP0_Insert
Next I
meYBDFCMP0.BDFCMPSTAT = mBDFCMPSTAT_Ligne & "3": GoTo YBDFCMP0_Insert

'------------------------------------------
YBDFCMP0_Insert:
'------------------------------------------
V = sqlYBDFCMP0_Insert(meYBDFCMP0)
If Not IsNull(V) Then GoTo Error_MsgBox

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
    
    YBDFCMP0_Add_ZCHGOPE0 = V
    '$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function
Public Function YBDFCMP0_Add_ZENCREG0()
Dim V, X As String, xSql As String, xWhere As String
Dim Nb As Long
Dim mMsgBox As String
Dim wDev As String, curDev As Currency, curEur As Currency, wAmj As Long
Dim wBDFCMPNAT As String
Dim blnEmis As Boolean, mBDFCMPSTAT_Ligne As String
Dim wCours As Double

On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "YBDFCMP0_Add_ZENCREG0"
'-------------------------------------------------------

YBDFCMP0_Add_ZENCREG0 = Null
mMsgBox = xYBDFCMP0.BDFCMPOPE & " " & xYBDFCMP0.BDFCMPNAT & " " & xYBDFCMP0.BDFCMPDOS
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

Call lstErr_AddItem(lstErr, cmdContext, "Ajout MAD : " & mMsgBox): DoEvents
'________________________________________________________________________________
'------------------------------------------
meYBDFCMP0 = xYBDFCMP0

Set rsSabX = Nothing
xWhere = " where COMPTECOM = '" & meYBDFCMP0.BDFCMPXCR & "'"
            
xSql = "select PLANCOPRO from " & paramIBM_Library_SABSPE & ".YBIACPT0 " & xWhere
Set rsSabX = cnsab.Execute(xSql)
If Not rsSabX.EOF Then meYBDFCMP0.BDFCMPXCRN = rsSabX("PLANCOPRO")

xWhere = " where COMPTECOM = '" & meYBDFCMP0.BDFCMPXDB & "'"
            
xSql = "select PLANCOPRO from " & paramIBM_Library_SABSPE & ".YBIACPT0 " & xWhere
Set rsSabX = cnsab.Execute(xSql)
If Not rsSabX.EOF Then meYBDFCMP0.BDFCMPXDBN = rsSabX("PLANCOPRO")

'-------------------------------------------------------------------------------------------------
If meYBDFCMP0.BDFCMPXCRN = "INT" Then meYBDFCMP0.BDFCMPSTAT = "721": meYBDFCMP0.BDFCMPSTA = "I": GoTo YBDFCMP0_Insert
If meYBDFCMP0.BDFCMPXDBN = "INT" Then meYBDFCMP0.BDFCMPSTAT = "721": meYBDFCMP0.BDFCMPSTA = "I": GoTo YBDFCMP0_Insert

If mId$(meYBDFCMP0.BDFCMPXDBN, 1, 1) = "L" Then
    If meYBDFCMP0.BDFCMPXCRN = "CAV" Or meYBDFCMP0.BDFCMPXCRN = "ASD" Then meYBDFCMP0.BDFCMPSTAT = "211": GoTo YBDFCMP0_Insert
    If mId$(meYBDFCMP0.BDFCMPXCRN, 1, 1) = "L" Then meYBDFCMP0.BDFCMPSTAT = "211": GoTo YBDFCMP0_Insert
End If
If mId$(meYBDFCMP0.BDFCMPXCRN, 1, 1) = "L" Then
    If meYBDFCMP0.BDFCMPXDBN = "CAV" Or meYBDFCMP0.BDFCMPXDBN = "ASD" Then meYBDFCMP0.BDFCMPSTAT = "211": GoTo YBDFCMP0_Insert
    If mId$(meYBDFCMP0.BDFCMPXDBN, 1, 1) = "L" Then meYBDFCMP0.BDFCMPSTAT = "211": GoTo YBDFCMP0_Insert
End If

If meYBDFCMP0.BDFCMPXCRN = "NOS" Then
    mBDFCMPSTAT_Ligne = "24"
    mENCREGCL5 = mENCREGCL5_C: mENCREGBC5 = mENCREGBC5_C: mENCREGLC5 = mENCREGLC5_C

Else
    If meYBDFCMP0.BDFCMPXDBN = "NOS" Then
        mBDFCMPSTAT_Ligne = "25"
        mENCREGCL5 = mENCREGCL5_D: mENCREGBC5 = mENCREGBC5_D: mENCREGLC5 = mENCREGLC5_D
    Else
        mENCREGCL5 = "": mENCREGBC5 = "": mENCREGLC5 = ""
    End If
End If
    
If Len(mENCREGLC5) > 1 Then
    If Not IsNumeric(mId$(mENCREGLC5, 1, 2)) Then meYBDFCMP0.BDFCMPPAYS = mId$(mENCREGLC5, 1, 2)
End If
If Trim(mENCREGBC5) <> "" Then
    meYBDFCMP0.BDFCMPBBIC = mENCREGBC5
Else
    meYBDFCMP0.BDFCMPBBIC = mENCREGCL5
End If
If meYBDFCMP0.BDFCMPPAYS = "  " Then meYBDFCMP0.BDFCMPPAYS = mId$(meYBDFCMP0.BDFCMPBBIC, 5, 2)

If meYBDFCMP0.BDFCMPDEV <> "EUR" Then
    If meYBDFCMP0.BDFCMPCREG = "SWF" Or meYBDFCMP0.BDFCMPCREG = "TGT" Then
        meYBDFCMP0.BDFCMPSTAT = mBDFCMPSTAT_Ligne & "3": GoTo YBDFCMP0_Insert
    Else
        meYBDFCMP0.BDFCMPSTAT = "751": GoTo YBDFCMP0_Insert
    End If
End If

If meYBDFCMP0.BDFCMPDEV = "EUR" And meYBDFCMP0.BDFCMPPAYS = "FR" Then meYBDFCMP0.BDFCMPSTAT = "711": meYBDFCMP0.BDFCMPSTA = "I": GoTo YBDFCMP0_Insert

If meYBDFCMP0.BDFCMPPAYS = "  " Then meYBDFCMP0.BDFCMPSTAT = "741": GoTo YBDFCMP0_Insert
For I = 1 To arrPAYS_UE_Nb
    If meYBDFCMP0.BDFCMPPAYS = arrPAYS_UE(I) Then meYBDFCMP0.BDFCMPSTAT = mBDFCMPSTAT_Ligne & "2": GoTo YBDFCMP0_Insert
Next I
meYBDFCMP0.BDFCMPSTAT = mBDFCMPSTAT_Ligne & "3": GoTo YBDFCMP0_Insert

'------------------------------------------
YBDFCMP0_Insert:
'------------------------------------------
V = sqlYBDFCMP0_Insert(meYBDFCMP0)
If Not IsNull(V) Then GoTo Error_MsgBox

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
    
    YBDFCMP0_Add_ZENCREG0 = V
    '$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function

Public Function YBDFCMP0_Add_ZMOUVEA0(mMOUVEMMON As Currency)
Dim V, X As String, xSql As String, xWhere As String
Dim Nb As Long
Dim mMsgBox As String
Dim wDev As String, curDev As Currency, curEur As Currency, wAmj As Long
Dim wBDFCMPNAT As String
Dim blnEmis As Boolean, mBDFCMPSTAT_Ligne As String
Dim mCDOREGUTI As Integer, mCDOREGREG As Integer
Dim mCDOREGDEC As String, mCDOREGDES As String
Dim wCours As Double
Dim wLong As Long

On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "YBDFCMP0_Add_ZMOUVEA0"
'-------------------------------------------------------
YBDFCMP0_Add_ZMOUVEA0 = Null
mMsgBox = xYBDFCMP0.BDFCMPOPE & " " & xYBDFCMP0.BDFCMPDOS & " " & xYBDFCMP0.BDFCMPNAT
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

Call lstErr_AddItem(lstErr, cmdContext, "Ajout MAD : " & mMsgBox): DoEvents
'________________________________________________________________________________
meYBDFCMP0 = xYBDFCMP0

'------------------------------------------

Set rsSabX = Nothing
xWhere = " where COMPTECOM = '" & meYBDFCMP0.BDFCMPXCR & "'"
            
xSql = "select COMPTEDEV,PLANCOPRO from " & paramIBM_Library_SABSPE & ".YBIACPT0 " & xWhere
Set rsSabX = cnsab.Execute(xSql)
If rsSabX.EOF Then meYBDFCMP0.BDFCMPSTAT = "891": GoTo YBDFCMP0_Insert
meYBDFCMP0.BDFCMPXCRN = rsSabX("PLANCOPRO")

'________________________________________________________________________________
meYBDFCMP0.BDFCMPDEV = rsSabX("COMPTEDEV")

If meYBDFCMP0.BDFCMPDEV = "EUR" Then
    meYBDFCMP0.BDFCMPMONE = meYBDFCMP0.BDFCMPMON
Else
    wAmj = meYBDFCMP0.BDFCMPDOPE - 19000000
    wCours = rsZBASTAB0_Cours37(meYBDFCMP0.BDFCMPDEV, wAmj)
    If wCours > 0 Then curEur = curDev / wCours

End If
'________________________________________________________________________________
mCDOREGUTI = Val(meYBDFCMP0.BDFCMPNAT)
Set rsSabX = Nothing
xWhere = " where CDOREGDOS = " & meYBDFCMP0.BDFCMPDOS _
       & " and CDOREGUTI = " & mCDOREGUTI _
       & " and CDOREGCRD = 'C' and CDOREGMON = " & cur_P(mMOUVEMMON)
xSql = "select CDOREGMOD,CDOREGPAY,CDOREGCOM,CDOREGDEV,CDOREGDEC,CDOREGDES,CDOREGREG" _
     & " from " & paramIBM_Library_SAB & ".ZCDOREG0 " & xWhere
Set rsSabX = cnsab.Execute(xSql)

If rsSabX.EOF Then
    xWhere = " where CDOREGDOS = " & meYBDFCMP0.BDFCMPDOS _
           & " and CDOREGUTI = " & mCDOREGUTI _
           & " and CDOREGCRD = 'C'"
    xSql = "select CDOREGMOD,CDOREGPAY,CDOREGCOM,CDOREGDEV,CDOREGDEC,CDOREGDES,CDOREGREG" _
         & " from " & paramIBM_Library_SAB & ".ZCDOREG0 " & xWhere
    Set rsSabX = cnsab.Execute(xSql)
End If
If rsSabX.EOF Then meYBDFCMP0.BDFCMPSTAT = "881": GoTo YBDFCMP0_Insert
mCDOREGREG = rsSabX("CDOREGREG")
If wLong > 9 Then meYBDFCMP0.BDFCMPSTAT = "881": GoTo YBDFCMP0_Insert
Mid$(meYBDFCMP0.BDFCMPNAT, 3, 1) = mCDOREGREG
meYBDFCMP0.BDFCMPCREG = rsSabX("CDOREGMOD")
meYBDFCMP0.BDFCMPPAYS = rsSabX("CDOREGPAY")

If meYBDFCMP0.BDFCMPXCR <> rsSabX("CDOREGCOM") Then meYBDFCMP0.BDFCMPSTAT = "871": GoTo YBDFCMP0_Insert
If meYBDFCMP0.BDFCMPDEV <> rsSabX("CDOREGDEV") Then meYBDFCMP0.BDFCMPSTAT = "861": GoTo YBDFCMP0_Insert
'________________________________________________________________________________
mCDOREGDEC = rsSabX("CDOREGDEC")
mCDOREGDES = rsSabX("CDOREGDES")

xWhere = " where CDOSWIDOS = " & meYBDFCMP0.BDFCMPDOS _
       & " and CDOSWIUTI = " & mCDOREGUTI _
       & " and CDOSWIREG = " & mCDOREGREG _
       & " and CDOSWICOP = '" & meYBDFCMP0.BDFCMPOPE & "'"
xSql = "select CDOSWIIBE" _
     & " from " & paramIBM_Library_SAB & ".ZCDOSWI0 " & xWhere
Set rsSabX = cnsab.Execute(xSql)
If Not rsSabX.EOF Then
    X = mId$(rsSabX("CDOSWIIBE"), 1, 2)
    If X <> "  " Then meYBDFCMP0.BDFCMPPAYS = X
End If
'________________________________________________________________________________

If meYBDFCMP0.BDFCMPPAYS = "  " And mCDOREGDEC = "T" Then
    xWhere = " where CDOTIEETB = 1 and CDOTIETIE = '" & mCDOREGDES & "'"
                
    xSql = "select CDOTIEPAR" _
         & " from " & paramIBM_Library_SAB & ".ZCDOTIE0 " & xWhere
    Set rsSabX = cnsab.Execute(xSql)
    
    
    If rsSabX.EOF Then meYBDFCMP0.BDFCMPSTAT = "851": GoTo YBDFCMP0_Insert
    X = Trim(rsSabX("CDOTIEPAR"))
    If meYBDFCMP0.BDFCMPPAYS = "  " Then
        meYBDFCMP0.BDFCMPPAYS = X
    Else
        If meYBDFCMP0.BDFCMPPAYS <> X Then meYBDFCMP0.BDFCMPPAYS = X ''''meYBDFCMP0.BDFCMPSTAT = "841": GoTo YBDFCMP0_Insert
    End If
End If

'________________________________________________________________________________
If meYBDFCMP0.BDFCMPCREG = "INT" Then meYBDFCMP0.BDFCMPSTAT = "821": meYBDFCMP0.BDFCMPSTA = "I": GoTo YBDFCMP0_Insert
If meYBDFCMP0.BDFCMPDEV = "EUR" And meYBDFCMP0.BDFCMPPAYS = "FR" Then meYBDFCMP0.BDFCMPSTAT = "811": meYBDFCMP0.BDFCMPSTA = "I": GoTo YBDFCMP0_Insert

If meYBDFCMP0.BDFCMPXCRN = "CAV" Or meYBDFCMP0.BDFCMPXCRN = "ASD" Then meYBDFCMP0.BDFCMPSTAT = "211": GoTo YBDFCMP0_Insert
If meYBDFCMP0.BDFCMPDEV <> "EUR" Then meYBDFCMP0.BDFCMPSTAT = "243": GoTo YBDFCMP0_Insert
'-------------------------------------------------------------------------------------------------

If meYBDFCMP0.BDFCMPPAYS = "  " Then meYBDFCMP0.BDFCMPSTAT = "841": GoTo YBDFCMP0_Insert
If meYBDFCMP0.BDFCMPPAYS = "FR" Then meYBDFCMP0.BDFCMPSTAT = "241": GoTo YBDFCMP0_Insert
For I = 1 To arrPAYS_UE_Nb
    If meYBDFCMP0.BDFCMPPAYS = arrPAYS_UE(I) Then meYBDFCMP0.BDFCMPSTAT = "242": GoTo YBDFCMP0_Insert
Next I
meYBDFCMP0.BDFCMPSTAT = "243": GoTo YBDFCMP0_Insert

'------------------------------------------
YBDFCMP0_Insert:
'------------------------------------------
V = sqlYBDFCMP0_Insert(meYBDFCMP0)
If Not IsNull(V) Then
    Mid$(meYBDFCMP0.BDFCMPNAT, 3, 1) = "*"
    V = sqlYBDFCMP0_Insert(meYBDFCMP0)
End If
If Not IsNull(V) Then GoTo Error_MsgBox

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
    
    YBDFCMP0_Add_ZMOUVEA0 = V
    '$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function


Public Function cmdUpdate_Ok_Transaction()
Dim V, X As String, xSql As String
Dim Nb As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdUpdate_Ok_Transaction"
'-------------------------------------------------------
cmdUpdate_Ok_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
V = sqlYBDFCMP0_Update(newYBDFCMP0, oldYBDFCMP0)
If Not IsNull(V) Then GoTo Error_MsgBox

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
    
    cmdUpdate_Ok_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function


Public Sub lstSelect_Load_5()
cmdSelect_Ok_Caption = "Importer les TRF de SAB=> BDFCMP "
'cmdSelect_Ok.BackColor = &HC0FFC0
txtSelect_AmjMin.Visible = True
txtSelect_AmjMin.Enabled = True
txtSelect_AmjMax.Visible = True
txtSelect_AmjMax.Enabled = True
Call DTPicker_Set(txtSelect_AmjMax, YBIATAB0_DATE_CAL_AP1)
Call DTPicker_Set(txtSelect_AmjMin, mId$(YBIATAB0_DATE_CAL_AP1, 1, 4) & "0101")

cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
End Sub

Public Sub lstSelect_Load_6()
cmdSelect_Ok_Caption = "Importer les CDO de SAB=> BDFCMP "
'cmdSelect_Ok.BackColor = &HC0FFC0
txtSelect_AmjMin.Visible = True
txtSelect_AmjMin.Enabled = True
txtSelect_AmjMax.Visible = True
txtSelect_AmjMax.Enabled = True
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
Call DTPicker_Set(txtSelect_AmjMax, YBIATAB0_DATE_CAL_AP1)
Call DTPicker_Set(txtSelect_AmjMin, mId$(YBIATAB0_DATE_CAL_AP1, 1, 4) & "0101")

cmdSelect_Ok.Visible = True
End Sub

Public Sub lstSelect_Load_7()
cmdSelect_Ok_Caption = "Importer les RD% de SAB=> BDFCMP "
'cmdSelect_Ok.BackColor = &HC0FFC0
txtSelect_AmjMin.Visible = True
txtSelect_AmjMin.Enabled = True
txtSelect_AmjMax.Visible = True
txtSelect_AmjMax.Enabled = True
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
Call DTPicker_Set(txtSelect_AmjMax, YBIATAB0_DATE_CAL_AP1)
Call DTPicker_Set(txtSelect_AmjMin, mId$(YBIATAB0_DATE_CAL_AP1, 1, 4) & "0101")

cmdSelect_Ok.Visible = True
End Sub

Public Sub fraUpdate_Display()
Dim X As String, xWhere As String, xSql As String

Call lstErr_Clear(lstErr, cmdContext, ">Afficahge du détail dossier"): DoEvents
lblUpdate_BDFCMPDOS = "Opération " & xYBDFCMP0.BDFCMPSER & xYBDFCMP0.BDFCMPSSE
lblUpdate_BDFCMPDOS.ForeColor = vbMagenta
'fraUpdate.ForeColor = vbMagenta
txtUpdate_BDFCMPOPE = xYBDFCMP0.BDFCMPOPE
txtUpdate_BDFCMPNAT = xYBDFCMP0.BDFCMPNAT
txtUpdate_BDFCMPDOS = xYBDFCMP0.BDFCMPDOS
txtUpdate_BDFCMPMON = Format(xYBDFCMP0.BDFCMPMON, "### ### ### ##0.00")
txtUpdate_BDFCMPDEV = xYBDFCMP0.BDFCMPDEV
txtUpdate_BDFCMPMONE = Format(xYBDFCMP0.BDFCMPMONE, "### ### ### ##0.00")
Call DTPicker_Set(txtUpdate_BDFCMPDCRE, CStr(xYBDFCMP0.BDFCMPDCRE))
Call DTPicker_Set(txtUpdate_BDFCMPDOPE, CStr(xYBDFCMP0.BDFCMPDOPE))
txtUpdate_BDFCMPCREG = Trim(xYBDFCMP0.BDFCMPCREG)
txtUpdate_BDFCMPXCR = Trim(xYBDFCMP0.BDFCMPXCR)
txtUpdate_BDFCMPXCRN = xYBDFCMP0.BDFCMPXCRN
Set rsSabX = Nothing
xWhere = " where COMPTECOM = '" & xYBDFCMP0.BDFCMPXCR & "'"
xSql = "select COMPTEINT from " & paramIBM_Library_SAB & ".ZCOMPTE0 " & xWhere
Set rsSabX = cnsab.Execute(xSql)
If Not rsSabX.EOF Then libUpdate_BDFCMPXCR = rsSabX("COMPTEINT")

txtUpdate_BDFCMPXDB = Trim(xYBDFCMP0.BDFCMPXDB)
txtUpdate_BDFCMPXDBN = xYBDFCMP0.BDFCMPXDBN
xWhere = " where COMPTECOM = '" & xYBDFCMP0.BDFCMPXDB & "'"
xSql = "select COMPTEINT from " & paramIBM_Library_SAB & ".ZCOMPTE0 " & xWhere
Set rsSabX = cnsab.Execute(xSql)
If Not rsSabX.EOF Then libUpdate_BDFCMPXDB = rsSabX("COMPTEINT")

txtUpdate_BDFCMPBBIC = Trim(xYBDFCMP0.BDFCMPBBIC)
txtUpdate_BDFCMPPAYS = Trim(xYBDFCMP0.BDFCMPPAYS)
txtUpdate_BDFCMPSTA = Trim(xYBDFCMP0.BDFCMPSTA)
txtUpdate_BDFCMPUSR = Trim(xYBDFCMP0.BDFCMPUSR)
txtUpdate_BDFCMPMTk = xYBDFCMP0.BDFCMPMTK & " " & xYBDFCMP0.BDFCMPROUT & " " & xYBDFCMP0.BDFCMPSABK
txtUpdate_BDFCMP50PI = xYBDFCMP0.BDFCMP50PI
txtUpdate_BDFCMP59PI = xYBDFCMP0.BDFCMP59PI


cboUpdate_BDFCMPSTAT.BackColor = txtUsr.BackColor
cbo_Scan xYBDFCMP0.BDFCMPSTAT, cboUpdate_BDFCMPSTAT
cboUpdate_BDFCMP2008.BackColor = txtUsr.BackColor
cbo_Scan xYBDFCMP0.BDFCMP2008, cboUpdate_BDFCMP2008


fraUpdate.Visible = True
fraUpdate_A.Enabled = False
fraUpdate_B.Enabled = BIA_BDFCMP_Aut.Saisir
txtUpdate_BDFCMPSTA.Enabled = BIA_BDFCMP_Aut.Saisir

End Sub

Private Sub txtSelect_BDFCMPOPE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtSelect_BDFCMPNAT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtSelect_BDFCMPSTAT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtUpdate_BDFCMPSTA_GotFocus()
Call txt_GotFocus(txtUpdate_BDFCMPSTA)

End Sub

Private Sub txtUpdate_BDFCMPSTA_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtUpdate_BDFCMPSTA_LostFocus()
Call txt_LostFocus(txtUpdate_BDFCMPSTA)

End Sub



Public Sub cmdSelect_SQL_2_Display(lBDFCMPSTAT As String, lBDFCMPMONE As Currency, lNb As Long)
Dim K As Integer

    cbo_Scan lBDFCMPSTAT, cboUpdate_BDFCMPSTAT
    fgStatistiques.Rows = fgStatistiques.Rows + 1
    fgStatistiques.Row = fgStatistiques.Rows - 1 'Val(mId$(lBDFCMPSTAT, 2, 1)) + 1
    fgStatistiques.Col = 0
    If cboUpdate_BDFCMPSTAT.ListIndex >= 0 Then fgStatistiques.Text = cboUpdate_BDFCMPSTAT
K = Val(mId$(lBDFCMPSTAT, 3, 1)) * 2 - 1
If K > 0 Then
    fgStatistiques.Col = K
Else
    fgStatistiques.Col = 1
End If

fgStatistiques.Text = Format$(lNb, "### ### ##0")
fgStatistiques.Col = fgStatistiques.Col + 1
fgStatistiques.Text = Format$(lBDFCMPMONE, "### ### ### ###.00")

arrYBDFCMP0(fgStatistiques.Row).BDFCMPSTAT = lBDFCMPSTAT
arrYBDFCMP0(fgStatistiques.Row).BDFCMPMON = lBDFCMPMONE
arrYBDFCMP0(fgStatistiques.Row).BDFCMPDOS = lNb

End Sub
Public Sub cmdSelect_SQL_4_Display(lBDFCMP2008 As String, lBDFCMPMONE As Currency, lNb As Long)
Dim K As Integer

K = Val(mId$(lBDFCMP2008, 4, 1)) * 2 - 1
If K > 0 Then
    fgStatistiques.Col = K
Else
    fgStatistiques.Col = 1
End If

fgStatistiques.Text = Format$(lNb, "### ### ##0")
fgStatistiques.Col = fgStatistiques.Col + 1
fgStatistiques.Text = Format$(lBDFCMPMONE, "### ### ### ###.00")

arrYBDFCMP0(fgStatistiques.Row).BDFCMP2008 = lBDFCMP2008
arrYBDFCMP0(fgStatistiques.Row).BDFCMPMON = lBDFCMPMONE
arrYBDFCMP0(fgStatistiques.Row).BDFCMPDOS = lNb

End Sub

Public Sub cmdSelect_SQL_3_Display(lBDFCMPOPE As String, lBDFCMPDEV As String, lBDFCMPSTAT As String, lBDFCMPMON As Currency, lNb As Long)
Dim K As Integer

    fgStatistiques.Rows = fgStatistiques.Rows + 1
    fgStatistiques.Row = fgStatistiques.Rows - 1 'Val(mId$(wBDFCMPSTAT, 2, 1)) + 1
    fgStatistiques.Col = 0
    fgStatistiques.Text = lBDFCMPOPE & " " & lBDFCMPDEV & " " & lBDFCMPSTAT
'End If
K = Val(mId$(lBDFCMPSTAT, 3, 1)) * 2 - 1
If K > 0 Then
    fgStatistiques.Col = K
Else
    fgStatistiques.Col = 1
End If

fgStatistiques.Text = Format$(lNb, "### ### ##0")
fgStatistiques.Col = fgStatistiques.Col + 1
fgStatistiques.Text = Format$(lBDFCMPMON, "### ### ### ###.00")

arrYBDFCMP0(fgStatistiques.Row).BDFCMPOPE = lBDFCMPOPE
arrYBDFCMP0(fgStatistiques.Row).BDFCMPDEV = lBDFCMPDEV
arrYBDFCMP0(fgStatistiques.Row).BDFCMPSTAT = lBDFCMPSTAT
arrYBDFCMP0(fgStatistiques.Row).BDFCMPMON = lBDFCMPMON
arrYBDFCMP0(fgStatistiques.Row).BDFCMPDOS = lNb

End Sub

Public Sub cmdPrint_Ok_3()
Dim K As Long, X As String
Dim wIndex As Integer
Dim sNb As Long, sMON As Currency
sNb = 0
sMON = 0
xYBDFCMP0.BDFCMPOPE = "": xYBDFCMP0.BDFCMPDEV = ""

prtBDF_CMP_Open 3, "BDF CMP: Comptage par Type d'opération / Devise / code Stat"
For K = 1 To fgStatistiques.Rows - 1
        If xYBDFCMP0.BDFCMPOPE <> arrYBDFCMP0(K).BDFCMPOPE _
        Or xYBDFCMP0.BDFCMPDEV <> arrYBDFCMP0(K).BDFCMPDEV Then
            If sNb > 0 Then cmdPrint_Ok_3_Total xYBDFCMP0.BDFCMPOPE, xYBDFCMP0.BDFCMPDEV, sNb, sMON
        End If
        xYBDFCMP0 = arrYBDFCMP0(K)
        
        prtBDF_CMP_NewLine 3
        XPrt.CurrentX = prtMinX + 50: XPrt.Print xYBDFCMP0.BDFCMPOPE;
        XPrt.CurrentX = prtMinX + 450: XPrt.Print xYBDFCMP0.BDFCMPDEV;
        XPrt.CurrentX = prtMinX + 850: XPrt.Print xYBDFCMP0.BDFCMPSTAT;
        X = Format$(xYBDFCMP0.BDFCMPDOS, " ### ### ##0")
        XPrt.CurrentX = prtMinX + 2000 - XPrt.TextWidth(X)
        XPrt.Print X;
        X = Format$(Abs(xYBDFCMP0.BDFCMPMON), "### ### ### ###.00")
        Select Case Val(mId$(xYBDFCMP0.BDFCMPSTAT, 3, 1))
         Case 2: XPrt.CurrentX = prtMinX + 8000 - XPrt.TextWidth(X)
         Case 3: XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
         Case Else: XPrt.CurrentX = prtMinX + 5000 - XPrt.TextWidth(X)
        End Select
        XPrt.Print X;
        sNb = sNb + xYBDFCMP0.BDFCMPDOS
        sMON = sMON + xYBDFCMP0.BDFCMPMON
Next K
If sNb > 0 Then cmdPrint_Ok_3_Total xYBDFCMP0.BDFCMPOPE, xYBDFCMP0.BDFCMPDEV, sNb, sMON
prtBDF_CMP_Close 3



End Sub


Public Sub cmdPrint_Ok_2()
Dim K As Long, X As String
Dim wIndex As Integer
Dim sNb As Long, sMON As Currency
sNb = 0
sMON = 0
xYBDFCMP0.BDFCMPOPE = "": xYBDFCMP0.BDFCMPDEV = ""

prtBDF_CMP_Open 2, "BDF CMP: Déclaration cartographie des moyens de paiement"
For K = 1 To fgStatistiques.Rows - 1
        xYBDFCMP0 = arrYBDFCMP0(K)
        
        prtBDF_CMP_NewLine 2
        cbo_Scan xYBDFCMP0.BDFCMPSTAT, cboUpdate_BDFCMPSTAT
        If cboUpdate_BDFCMPSTAT.ListIndex >= 0 Then

            XPrt.CurrentX = prtMinX + 50: XPrt.Print cboUpdate_BDFCMPSTAT;
        End If
        X = Format$(xYBDFCMP0.BDFCMPDOS, " ### ### ##0")
        XPrt.CurrentX = prtMinX + 4000 - XPrt.TextWidth(X)
        XPrt.Print X;
        X = Format$(Abs(xYBDFCMP0.BDFCMPMON), "### ### ### ###.00")
        Select Case Val(mId$(xYBDFCMP0.BDFCMPSTAT, 3, 1))
         Case 2: XPrt.CurrentX = prtMinX + 9000 - XPrt.TextWidth(X)
         Case 3: XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
         Case Else: XPrt.CurrentX = prtMinX + 7000 - XPrt.TextWidth(X)
        End Select
        XPrt.Print X;
        sNb = sNb + xYBDFCMP0.BDFCMPDOS
        sMON = sMON + xYBDFCMP0.BDFCMPMON
Next K
prtBDF_CMP_Close 2



End Sub

Public Sub cmdPrint_Ok_4()
Dim K As Long, X As String
Dim wIndex As Integer
Dim sNb As Long, sMON As Currency
sNb = 0
sMON = 0
xYBDFCMP0.BDFCMPOPE = "": xYBDFCMP0.BDFCMPDEV = ""

prtBDF_CMP_Open 4, "BDF CMP: Déclaration cartographie des moyens de paiement"
For K = 1 To fgStatistiques.Rows - 1
    
    prtBDF_CMP_NewLine 4
    fgStatistiques.Row = K
    fgStatistiques.Col = 0:
    XPrt.CurrentX = prtMinX + 50: XPrt.Print fgStatistiques.Text;
    fgStatistiques.Col = 1: X = Format$(num_CDec(fgStatistiques.Text), "### ### ###")
    XPrt.CurrentX = prtMinX + 2500 - XPrt.TextWidth(X): XPrt.Print X;
    fgStatistiques.Col = 2: X = Format$(num_CDec(fgStatistiques.Text), "### ### ### ###.##")
    XPrt.CurrentX = prtMinX + 6000 - XPrt.TextWidth(X): XPrt.Print X;
    fgStatistiques.Col = 3: X = Format$(num_CDec(fgStatistiques.Text), "### ### ###")
    XPrt.CurrentX = prtMinX + 7500 - XPrt.TextWidth(X): XPrt.Print X;
    fgStatistiques.Col = 4:  X = Format$(num_CDec(fgStatistiques.Text), "### ### ### ###.##")
    XPrt.CurrentX = prtMinX + 11000 - XPrt.TextWidth(X): XPrt.Print X
    fgStatistiques.Col = 5: X = Format$(num_CDec(fgStatistiques.Text), "### ### ###")
    XPrt.CurrentX = prtMinX + 12500 - XPrt.TextWidth(X): XPrt.Print X;
    fgStatistiques.Col = 6: X = Format$(num_CDec(fgStatistiques.Text), "### ### ### ###.##")
    XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X): XPrt.Print X;
Next K
prtBDF_CMP_Close 4



End Sub


Public Sub cmdPrint_Ok_3_Total(lOPE As String, lDEV As String, sNb As Long, sMON As Currency)
prtBDF_CMP_NewLine 3
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight, " ", 240)
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 50: XPrt.Print lOPE;
XPrt.CurrentX = prtMinX + 450: XPrt.Print lDEV;

X = Format$(sNb, " ### ### ##0")
XPrt.CurrentX = prtMinX + 2000 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(Abs(sMON), "### ### ### ###.00")
XPrt.CurrentX = prtMinX + 5000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.FontBold = False
sNb = 0: sMON = 0
End Sub

Public Sub YBDFCMP0_V2008_Export_Add(Nb As Long, nbk As Long, mtk As Currency)
'____________________________________________________________________________________

Dim iBackColor  As Integer
If nbk > 0 Then
        Nb = Nb + 1
        If Nb Mod 10 = 0 Then Call lstErr_ChangeLastItem(lstErr, cmdContext, "Export en cours : " & Nb & " enregistrements"): DoEvents

        wsExcel.Cells(Nb, 1) = oldYBDFCMP0.BDFCMP2008
        If mId$(oldYBDFCMP0.BDFCMP2008, 1, 1) = "T" Or mId$(oldYBDFCMP0.BDFCMP2008, 1, 1) = "X" Then
            iBackColor = 15
        Else
            Select Case mId$(oldYBDFCMP0.BDFCMP2008, 2, 1)
                Case "A": iBackColor = 6
                Case "B": iBackColor = 4
                Case "C": iBackColor = 8
                Case "S": iBackColor = 4
                Case "T": iBackColor = 7
                Case Else: iBackColor = 3
                
            End Select
        End If
        wsExcel.Cells(Nb, 1).Interior.ColorIndex = iBackColor

        wsExcel.Cells(Nb, 2) = oldYBDFCMP0.BDFCMPMTK
        wsExcel.Cells(Nb, 3) = oldYBDFCMP0.BDFCMPROUT
        wsExcel.Cells(Nb, 4) = oldYBDFCMP0.BDFCMPXCRN
        wsExcel.Cells(Nb, 5) = oldYBDFCMP0.BDFCMPXDBN
        wsExcel.Cells(Nb, 6) = oldYBDFCMP0.BDFCMPCREG
        wsExcel.Cells(Nb, 7) = oldYBDFCMP0.BDFCMPSER
        wsExcel.Cells(Nb, 8) = oldYBDFCMP0.BDFCMPSSE
        wsExcel.Cells(Nb, 9) = oldYBDFCMP0.BDFCMPOPE
        wsExcel.Cells(Nb, 10) = oldYBDFCMP0.BDFCMPNAT
        wsExcel.Cells(Nb, 11) = oldYBDFCMP0.BDFCMPDOS
        wsExcel.Cells(Nb, 12) = nbk
        wsExcel.Cells(Nb, 13) = mtk
End If

End Sub
