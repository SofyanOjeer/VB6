VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmYSWIDOS0 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB : SWIFT émis"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13575
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "YSWIDOS0.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10530
   ScaleWidth      =   13575
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   0
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9852
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   17383
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "statistiques SWIFT émis"
      TabPicture(0)   =   "YSWIDOS0.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "."
      TabPicture(1)   =   "YSWIDOS0.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstW"
      Tab(1).ControlCount=   1
      Begin VB.ListBox lstW 
         BackColor       =   &H00E0F0FF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   5100
         Left            =   -68160
         TabIndex        =   13
         Top             =   1680
         Visible         =   0   'False
         Width           =   5172
      End
      Begin VB.Frame fraTab0 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9420
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   13296
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   6540
            Left            =   1080
            TabIndex        =   12
            Top             =   2760
            Visible         =   0   'False
            Width           =   12252
            _ExtentX        =   21616
            _ExtentY        =   11536
            _Version        =   393216
            Cols            =   15
            FixedCols       =   0
            BackColor       =   15794175
            ForeColor       =   16711680
            BackColorFixed  =   8421504
            ForeColorFixed  =   -2147483633
            BackColorBkg    =   -2147483633
            FormatString    =   $"YSWIDOS0.frx":0342
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
            Left            =   9480
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   2400
            Width           =   2172
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   11760
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   2280
            Width           =   1212
         End
         Begin VB.Frame fraSelect_Options 
            BackColor       =   &H00F0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2052
            Left            =   0
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   13152
            Begin VB.ComboBox cboSelect_SWIDOS_5 
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
               Height          =   324
               Left            =   11400
               Style           =   2  'Dropdown List
               TabIndex        =   55
               Top             =   1600
               Width           =   1572
            End
            Begin VB.ComboBox cboSelect_SWIDOS_4 
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
               Height          =   324
               Left            =   11400
               Style           =   2  'Dropdown List
               TabIndex        =   45
               Top             =   1300
               Width           =   1572
            End
            Begin VB.ComboBox cboSelect_SWIDOS_3 
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
               Height          =   324
               Left            =   11400
               Style           =   2  'Dropdown List
               TabIndex        =   44
               Top             =   1000
               Width           =   1572
            End
            Begin VB.ComboBox cboSelect_SWIDOS_2 
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
               Height          =   324
               Left            =   11400
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   700
               Width           =   1572
            End
            Begin VB.ComboBox cboSelect_SWIDOS_1 
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
               Height          =   324
               Left            =   11400
               Style           =   2  'Dropdown List
               TabIndex        =   42
               Top             =   400
               Width           =   1572
            End
            Begin VB.Frame fraSelect_Options_1 
               BackColor       =   &H00D0F0FF&
               BorderStyle     =   0  'None
               Height          =   1812
               Left            =   120
               TabIndex        =   9
               Top             =   120
               Width           =   10932
               Begin VB.ComboBox cboSelect_SWIDOSBPIZ_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   9240
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   64
                  Top             =   1440
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWIDOSDPIZ_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   9240
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   63
                  Top             =   480
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWIDOSBPIZ 
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
                  Left            =   9840
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   62
                  Top             =   1440
                  Width           =   972
               End
               Begin VB.ComboBox cboSelect_SWIDOSDPIZ 
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
                  Left            =   9840
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   61
                  Top             =   480
                  Width           =   972
               End
               Begin VB.TextBox txtSElect_SWIDOS21 
                  Height          =   288
                  Left            =   600
                  TabIndex        =   60
                  Top             =   1500
                  Width           =   1812
               End
               Begin VB.TextBox txtSelect_SWIDOS20 
                  Height          =   288
                  Left            =   600
                  TabIndex        =   59
                  Top             =   1200
                  Width           =   1812
               End
               Begin VB.ComboBox cboSelect_SWIDOSMTK2 
                  Height          =   312
                  Left            =   3960
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   53
                  Top             =   480
                  Width           =   936
               End
               Begin VB.ComboBox cboSelect_SWIDOS57A 
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
                  Left            =   6600
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   41
                  Top             =   1440
                  Width           =   1812
               End
               Begin VB.ComboBox cboSelect_SWIDOS52A 
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
                  Left            =   6600
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   40
                  Top             =   1080
                  Width           =   1812
               End
               Begin VB.ComboBox cboSelect_SWIDOSRCV 
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
                  Left            =   6600
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   39
                  Top             =   120
                  Width           =   1812
               End
               Begin VB.ComboBox cboSelect_SWIDOSDEV 
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
                  Left            =   1200
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   38
                  Top             =   840
                  Width           =   1212
               End
               Begin VB.ComboBox cboSelect_SWIDOS59PI 
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
                  Left            =   9840
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   37
                  Top             =   1080
                  Width           =   972
               End
               Begin VB.ComboBox cboSelect_SWIDOS50PI 
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
                  Left            =   9840
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   36
                  Top             =   120
                  Width           =   972
               End
               Begin VB.ComboBox cboSelect_SWIDOSROUT 
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
                  Left            =   6600
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   35
                  Top             =   480
                  Width           =   972
               End
               Begin VB.ComboBox cboSelect_SWIDOSDEV_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   600
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   34
                  Top             =   840
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWIDOS59PI_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   9240
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   33
                  Top             =   1080
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWIDOS50PI_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   9240
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   32
                  Top             =   120
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWIDOSROUT_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   6000
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   31
                  Top             =   480
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWIDOS57A_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   6000
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   30
                  Top             =   1440
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWIDOS52A_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   6000
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   29
                  Top             =   1080
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWIDOSRCV_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   6000
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   28
                  Top             =   120
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWIDOSSSE 
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
                  Left            =   3960
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   20
                  Top             =   1440
                  Width           =   972
               End
               Begin VB.ComboBox cboSelect_SWIDOSSSE_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   3360
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   19
                  Top             =   1440
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWIDOSOPEC_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   3360
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   17
                  Top             =   960
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWIDOSMTK_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   3360
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   16
                  Top             =   120
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWIDOSOPEC 
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
                  Left            =   3960
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   14
                  Top             =   960
                  Width           =   972
               End
               Begin VB.ComboBox cboSelect_SWIDOSMTK 
                  Height          =   312
                  Left            =   3960
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   11
                  Top             =   120
                  Width           =   936
               End
               Begin MSComCtl2.DTPicker txtSelect_SWIDOSAMJ_Min 
                  Height          =   300
                  Left            =   1200
                  TabIndex        =   51
                  Top             =   120
                  Width           =   1212
                  _ExtentX        =   2143
                  _ExtentY        =   529
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
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
                  Format          =   37421059
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin MSComCtl2.DTPicker txtSelect_SWIDOSAMJ_Max 
                  Height          =   300
                  Left            =   1200
                  TabIndex        =   52
                  Top             =   480
                  Width           =   1212
                  _ExtentX        =   2143
                  _ExtentY        =   529
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
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
                  Format          =   37421059
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin VB.Label lblSelect_SWIDOSBPIZ 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "FR UE **"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   8520
                  TabIndex        =   66
                  Top             =   1500
                  Width           =   732
               End
               Begin VB.Label lblSelect_SWIDOSDPIZ 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "FR UE **"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   8520
                  TabIndex        =   65
                  Top             =   550
                  Width           =   732
               End
               Begin VB.Label lblSelect_SWIDOS21 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   ":21:"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   120
                  TabIndex        =   58
                  Top             =   1560
                  Width           =   372
               End
               Begin VB.Label lblSelect_SWIDOS20 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   ":20:"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   120
                  TabIndex        =   57
                  Top             =   1200
                  Width           =   372
               End
               Begin VB.Label lblSelect_SWIDOSAMJ 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "Date d'émission"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   0
                  TabIndex        =   50
                  Top             =   120
                  Width           =   1212
               End
               Begin VB.Label lblSelect_SWIDOSDEV 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "Devise"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   0
                  TabIndex        =   27
                  Top             =   840
                  Width           =   612
               End
               Begin VB.Label lblSelect_SWIDOS69PI 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "Pays BEN"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   8520
                  TabIndex        =   26
                  Top             =   1150
                  Width           =   732
               End
               Begin VB.Label lblSelect_SWIDOS50PI 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "Pays DO"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   8520
                  TabIndex        =   25
                  Top             =   200
                  Width           =   732
               End
               Begin VB.Label lblSelect_SWIDOSROUT 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "Routage"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   5280
                  TabIndex        =   24
                  Top             =   600
                  Width           =   732
               End
               Begin VB.Label lblSelect_SWIDOS57A 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "BIC BEN"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   5280
                  TabIndex        =   23
                  Top             =   1500
                  Width           =   732
               End
               Begin VB.Label lblSelect_SWIDOS52A 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "BIC DO"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   5280
                  TabIndex        =   22
                  Top             =   1080
                  Width           =   612
               End
               Begin VB.Label lblSelect_SWIDOSRCV 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "BIC RCV"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   5280
                  TabIndex        =   21
                  Top             =   240
                  Width           =   732
               End
               Begin VB.Label lblSelect_SWIDOSSSE 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "Service"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   2760
                  TabIndex        =   18
                  Top             =   1500
                  Width           =   612
               End
               Begin VB.Label lblSelect_SWIDOSOPEC 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "Code opé"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   2640
                  TabIndex        =   15
                  Top             =   996
                  Width           =   732
               End
               Begin VB.Label lblSelect_SWIDOSMTK 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "MTxxx"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   2760
                  TabIndex        =   10
                  Top             =   120
                  Width           =   612
               End
            End
            Begin VB.Label libSelect_SWIDOS 
               Alignment       =   2  'Center
               BackColor       =   &H00F0FFFF&
               Caption         =   "Tri- Ventilation "
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   11400
               TabIndex        =   56
               Top             =   120
               Width           =   1452
            End
            Begin VB.Label libSelect_SWIDOS_5 
               BackColor       =   &H00F0FFFF&
               Caption         =   "5 - "
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   11160
               TabIndex        =   54
               Top             =   1650
               Width           =   252
            End
            Begin VB.Label libSelect_SWIDOS_4 
               BackColor       =   &H00F0FFFF&
               Caption         =   "4 - "
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   11160
               TabIndex        =   49
               Top             =   1350
               Width           =   252
            End
            Begin VB.Label libSelect_SWIDOS_3 
               BackColor       =   &H00F0FFFF&
               Caption         =   "3 - "
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   11160
               TabIndex        =   48
               Top             =   1050
               Width           =   252
            End
            Begin VB.Label libSelect_SWIDOS_2 
               BackColor       =   &H00F0FFFF&
               Caption         =   "2 - "
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   11160
               TabIndex        =   47
               Top             =   750
               Width           =   252
            End
            Begin VB.Label libSelect_SWIDOS_1 
               BackColor       =   &H00F0FFFF&
               Caption         =   "1 - "
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   11160
               TabIndex        =   46
               Top             =   450
               Width           =   252
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7068
            Left            =   0
            TabIndex        =   5
            Top             =   2280
            Visible         =   0   'False
            Width           =   8352
            _ExtentX        =   14737
            _ExtentY        =   12462
            _Version        =   393216
            Rows            =   1
            Cols            =   7
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   -2147483633
            ForeColor       =   12582912
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483637
            BackColorSel    =   12648384
            BackColorBkg    =   -2147483633
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   "< 1    |< 2       |> Nb               |> Montant           "
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
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   13080
      Picture         =   "YSWIDOS0.frx":041C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuContextX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "mnuPrint"
      Visible         =   0   'False
      Begin VB.Menu mnuExportation 
         Caption         =   "Exportation .xlsx"
      End
   End
End
Attribute VB_Name = "frmYSWIDOS0"
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
Dim YSWIDOS0_Aut As typeAuthorization
Dim blnAuto As Boolean, blnError As Boolean
Dim cmdSelect_SQL_K As String

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
Dim xYSWIDOS0 As typeYSWIDOS0, newYSWIDOS0 As typeYSWIDOS0, oldYSWIDOS0 As typeYSWIDOS0
Dim arrYSWIDOS0() As typeYSWIDOS0, arrYSWIDOS0_Nb As Long, arrYSWIDOS0_Max As Long, arrYSWIDOS0_Index As Long


Dim fgDetail_FormatString As String, fgDetail_K As Integer
Dim fgDetail_RowDisplay As Integer, fgDetail_RowClick As Integer, fgDetail_ColClick As Integer
Dim fgDetail_ColorClick As Long, fgDetail_ColorDisplay As Long
Dim fgDetail_Sort1 As Integer, fgDetail_Sort2 As Integer
Dim fgDetail_SortAD As Integer, fgDetail_Sort1_Old As Integer
Dim fgDetail_arrIndex As Integer
Dim blnfgDetail_DisplayLine As Boolean


Dim arrSWIDOS_Field(12) As String, arrSWIDOS_Lib(12) As String, arrSWIDOS_Field_Nb As Integer
Dim arrSWIDOS_Group(12) As Integer, arrSWIDOS_Group_Nb As Integer
Dim mGroupBy As String
Dim xWhere_SQL As String
'______________________________________________________________________
Private Sub fgSelect_Display()
Dim wColor As Long

Dim I As Long, K As Integer
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect.Cols = arrSWIDOS_Group_Nb + 2
fgSelect_Reset

fgSelect.Rows = 1
fgSelect_FormatString = "<" & arrSWIDOS_Lib(arrSWIDOS_Group(1))
For K = 2 To arrSWIDOS_Group_Nb
    fgSelect_FormatString = fgSelect_FormatString & "|<" & arrSWIDOS_Lib(arrSWIDOS_Group(K))
Next K

fgSelect.FormatString = fgSelect_FormatString & "|>        Nombre |>                      Montant"
'fgSelect_FormatString
fgSelect.Row = 0

currentAction = "fgSelect_Display"

Do While Not rsSab.EOF
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine I, True

    rsSab.MoveNext
Loop
    

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYSWIDOS0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub arrYSWIDOS0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYSWIDOS0(101)
arrYSWIDOS0_Max = 100: arrYSWIDOS0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIDOS0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYSWIDOS0_GetBuffer(rsSab, xYSWIDOS0)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYSWIDOS0.fgselect_Display"
        '' Exit Sub
     Else
         arrYSWIDOS0_Nb = arrYSWIDOS0_Nb + 1
         If arrYSWIDOS0_Nb > arrYSWIDOS0_Max Then
             arrYSWIDOS0_Max = arrYSWIDOS0_Max + 100
             ReDim Preserve arrYSWIDOS0(arrYSWIDOS0_Max)
         End If
         
         arrYSWIDOS0(arrYSWIDOS0_Nb) = xYSWIDOS0
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

Public Sub cmdSelect_Reset()
If blnControl Then
    lstErr.Clear
    fgSelect.Visible = False
    fgDetail.Visible = False
    lstW.Visible = False
    cmdSelect_Ok.Visible = True
    cmdSelect_SQL_K = Trim(Mid$(cboSelect_SQL, 1, 2))
    Select Case cmdSelect_SQL_K
        Case "1":
            fraSelect_Options.Visible = True: fraSelect_Options_1.Visible = True
    End Select

End If

End Sub



Private Sub cmdSelect_SQL_1()
Dim V
Dim X As String, K As Integer
Dim xAnd As String, xSql As String
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdYSWIDOS0_SQL"
blnOk = False
Call DTPicker_Control(txtSelect_SWIDOSAMJ_Min, wAmjMin)
Call DTPicker_Control(txtSelect_SWIDOSAMJ_Max, wAmjMax)

xWhere_SQL = " Where SWIDOSDENV >= " & wAmjMin & " and SWIDOSDENV <= " & wAmjMax
If Trim(txtSelect_SWIDOS20) <> "" Then xWhere_SQL = xWhere_SQL & " and   SWIDOS20 like '%" & Trim(txtSelect_SWIDOS20) & "%'"
If Trim(txtSElect_SWIDOS21) <> "" Then xWhere_SQL = xWhere_SQL & " and   SWIDOS21 like '%" & Trim(txtSElect_SWIDOS21) & "%'"

If Trim(cboSelect_SWIDOSMTK_K) <> "" Then
    If Trim(cboSelect_SWIDOSMTK2) = "" Then
        xWhere_SQL = xWhere_SQL & " and   SWIDOSMTK " & Trim(cboSelect_SWIDOSMTK_K) & "'" & cboSelect_SWIDOSMTK & "'"
    Else
        xWhere_SQL = xWhere_SQL & " and  ( SWIDOSMTK " & Trim(cboSelect_SWIDOSMTK_K) & "'" & cboSelect_SWIDOSMTK & "' or   SWIDOSMTK " & Trim(cboSelect_SWIDOSMTK_K) & "'" & cboSelect_SWIDOSMTK2 & "')"
    End If
End If

If Trim(cboSelect_SWIDOSOPEC_K) <> "" Then xWhere_SQL = xWhere_SQL & " and   SWIDOSOPEC " & Trim(cboSelect_SWIDOSOPEC_K) & "'" & cboSelect_SWIDOSOPEC & "'"
 If Trim(cboSelect_SWIDOSDEV_K) <> "" Then xWhere_SQL = xWhere_SQL & " and   SWIDOSDEV " & Trim(cboSelect_SWIDOSDEV_K) & "'" & cboSelect_SWIDOSDEV & "'"
If Trim(cboSelect_SWIDOSSSE_K) <> "" Then xWhere_SQL = xWhere_SQL & " and   SWIDOSSSE " & Trim(cboSelect_SWIDOSSSE_K) & "'" & cboSelect_SWIDOSSSE & "'"
X = Trim(cboSelect_SWIDOSRCV_K)
Select Case X
    Case "":
    Case "=", "<>": xWhere_SQL = xWhere_SQL & " and   SWIDOSRCV " & Trim(cboSelect_SWIDOSRCV_K) & "'" & cboSelect_SWIDOSRCV & "'"
    Case "=4": xWhere_SQL = xWhere_SQL & " and   SWIDOSRCV like '" & Mid$(cboSelect_SWIDOSRCV, 1, 4) & "%'"
    Case "=6": xWhere_SQL = xWhere_SQL & " and   SWIDOSRCV like '" & Mid$(cboSelect_SWIDOSRCV, 1, 6) & "%'"
    Case "=8": xWhere_SQL = xWhere_SQL & " and   SWIDOSRCV like '" & Mid$(cboSelect_SWIDOSRCV, 1, 8) & "%'"
End Select

X = Trim(cboSelect_SWIDOS52A_K)
Select Case X
    Case "":
    Case "=", "<>": xWhere_SQL = xWhere_SQL & " and   SWIDOS52A " & Trim(cboSelect_SWIDOS52A_K) & "'" & cboSelect_SWIDOS52A & "'"
    Case "=4": xWhere_SQL = xWhere_SQL & " and   SWIDOS52A like '" & Mid$(cboSelect_SWIDOS52A, 1, 4) & "%'"
    Case "=6": xWhere_SQL = xWhere_SQL & " and   SWIDOS52A like '" & Mid$(cboSelect_SWIDOS52A, 1, 6) & "%'"
    Case "=8": xWhere_SQL = xWhere_SQL & " and   SWIDOS52A like '" & Mid$(cboSelect_SWIDOS52A, 1, 8) & "%'"
End Select

X = Trim(cboSelect_SWIDOS57A_K)
Select Case X
    Case "":
    Case "=", "<>": xWhere_SQL = xWhere_SQL & " and   SWIDOS57A " & Trim(cboSelect_SWIDOS57A_K) & "'" & cboSelect_SWIDOS57A & "'"
    Case "=4": xWhere_SQL = xWhere_SQL & " and   SWIDOS57A like '" & Mid$(cboSelect_SWIDOS57A, 1, 4) & "%'"
    Case "=6": xWhere_SQL = xWhere_SQL & " and   SWIDOS57A like '" & Mid$(cboSelect_SWIDOS57A, 1, 6) & "%'"
    Case "=8": xWhere_SQL = xWhere_SQL & " and   SWIDOS57A like '" & Mid$(cboSelect_SWIDOS57A, 1, 8) & "%'"
End Select

If Trim(cboSelect_SWIDOSROUT_K) <> "" Then xWhere_SQL = xWhere_SQL & " and   SWIDOSROUT " & Trim(cboSelect_SWIDOSROUT_K) & "'" & cboSelect_SWIDOSROUT & "'"
If Trim(cboSelect_SWIDOS50PI_K) <> "" Then xWhere_SQL = xWhere_SQL & " and   SWIDOS50PI " & Trim(cboSelect_SWIDOS50PI_K) & "'" & cboSelect_SWIDOS50PI & "'"
If Trim(cboSelect_SWIDOS59PI_K) <> "" Then xWhere_SQL = xWhere_SQL & " and   SWIDOS59PI " & Trim(cboSelect_SWIDOS59PI_K) & "'" & cboSelect_SWIDOS59PI & "'"
If Trim(cboSelect_SWIDOSDPIZ_K) <> "" Then xWhere_SQL = xWhere_SQL & " and   SWIDOSDPIZ " & Trim(cboSelect_SWIDOSDPIZ_K) & "'" & cboSelect_SWIDOSDPIZ & "'"
If Trim(cboSelect_SWIDOSBPIZ_K) <> "" Then xWhere_SQL = xWhere_SQL & " and   SWIDOSBPIZ " & Trim(cboSelect_SWIDOSBPIZ_K) & "'" & cboSelect_SWIDOSBPIZ & "'"
   
blnOk = False
arrSWIDOS_Group_Nb = 0
For K = 1 To arrSWIDOS_Field_Nb
    If Trim(cboSelect_SWIDOS_1) = arrSWIDOS_Lib(K) Then
        arrSWIDOS_Group_Nb = arrSWIDOS_Group_Nb + 1
        arrSWIDOS_Group(1) = K
        Exit For
    End If
Next K

If Trim(cboSelect_SWIDOS_2) = "" Then
    blnOk = True
Else
    For K = 1 To arrSWIDOS_Field_Nb
        If Trim(cboSelect_SWIDOS_2) = arrSWIDOS_Lib(K) Then
            arrSWIDOS_Group_Nb = arrSWIDOS_Group_Nb + 1
            arrSWIDOS_Group(arrSWIDOS_Group_Nb) = K
            Exit For
        End If
    Next K
End If
If Trim(cboSelect_SWIDOS_3) = "" Then
    blnOk = True
Else
    For K = 1 To arrSWIDOS_Field_Nb
        If Trim(cboSelect_SWIDOS_3) = arrSWIDOS_Lib(K) Then
            arrSWIDOS_Group_Nb = arrSWIDOS_Group_Nb + 1
            arrSWIDOS_Group(arrSWIDOS_Group_Nb) = K
            Exit For
        End If
    Next K
End If
If Trim(cboSelect_SWIDOS_4) = "" Then
    blnOk = True
Else
    For K = 1 To arrSWIDOS_Field_Nb
        If Trim(cboSelect_SWIDOS_4) = arrSWIDOS_Lib(K) Then
            arrSWIDOS_Group_Nb = arrSWIDOS_Group_Nb + 1
            arrSWIDOS_Group(arrSWIDOS_Group_Nb) = K
            Exit For
        End If
    Next K
End If
If Trim(cboSelect_SWIDOS_5) = "" Then
    blnOk = True
Else
    For K = 1 To arrSWIDOS_Field_Nb
        If Trim(cboSelect_SWIDOS_5) = arrSWIDOS_Lib(K) Then
            arrSWIDOS_Group_Nb = arrSWIDOS_Group_Nb + 1
            arrSWIDOS_Group(arrSWIDOS_Group_Nb) = K
            Exit For
        End If
    Next K
End If
   
mGroupBy = arrSWIDOS_Field(arrSWIDOS_Group(1))
For K = 2 To arrSWIDOS_Group_Nb
    mGroupBy = mGroupBy & " , " & arrSWIDOS_Field(arrSWIDOS_Group(K))
Next K
    
xSql = "select " & mGroupBy & " , count(*) , SUM(SWIDOSMON) from " & paramIBM_Library_SABSPE & ".YSWIDOS0" _
     & xWhere_SQL _
     & " group by  " & mGroupBy _
     & " order by " & mGroupBy
Set rsSab = cnsab.Execute(xSql)

fgSelect_Display


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub fgDetail_Display()
Dim wColor As Long
Dim xWhere As String, xSql As String

Dim K As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler
fgDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString
fgDetail.Row = 0

currentAction = "fgDetail_Display"

xWhere = ""
mGroupBy = arrSWIDOS_Field(arrSWIDOS_Group(1))
For K = 1 To arrSWIDOS_Group_Nb
    fgSelect.Col = K - 1
    xWhere = xWhere & " and " & arrSWIDOS_Field(arrSWIDOS_Group(K)) & " ='" & Trim(fgSelect.Text) & "'"
Next K

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIDOS0" _
     & xWhere_SQL & xWhere
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    V = rsYSWIDOS0_GetBuffer(rsSab, xYSWIDOS0)
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_DisplayLine
    
    rsSab.MoveNext
Loop

fgDetail.Visible = True


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long, blnYSWIDOS0 As Boolean)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim xSql As String, xCur As Currency
On Error Resume Next

For K = 0 To arrSWIDOS_Group_Nb - 1

    fgSelect.Col = K: fgSelect.Text = rsSab(K)
Next K
fgSelect.Col = arrSWIDOS_Group_Nb: fgSelect.Text = Format(rsSab(arrSWIDOS_Group_Nb), "##### ###")
xCur = rsSab(arrSWIDOS_Group_Nb + 1)
If xCur <> 0 Then fgSelect.Col = arrSWIDOS_Group_Nb + 1: fgSelect.Text = Format(xCur, "### ### ### ##0.00")


'fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
End Sub
Public Sub fgDetail_DisplayLine()
Dim K As Integer
Dim wColor As Long, wColor_Row As Long

On Error Resume Next
'wColor = vbBlue: wColor_Row = vbWhite
fgDetail.Col = 0: fgDetail.Text = xYSWIDOS0.SWIDOSSER & " " & xYSWIDOS0.SWIDOSSSE
fgDetail.Col = 1: fgDetail.Text = xYSWIDOS0.SWIDOSOPEC & " " & xYSWIDOS0.SWIDOSOPEN
fgDetail.Col = 2: fgDetail.Text = xYSWIDOS0.SWIDOSMTK

fgDetail.Col = 3: fgDetail.Text = xYSWIDOS0.SWIDOSROUT
fgDetail.Col = 4: fgDetail.Text = xYSWIDOS0.SWIDOSMON
fgDetail.Col = 5: fgDetail.Text = xYSWIDOS0.SWIDOSDEV
fgDetail.Col = 6: fgDetail.Text = dateImp10_S(xYSWIDOS0.SWIDOSDENV)
fgDetail.Col = 7: fgDetail.Text = xYSWIDOS0.SWIDOSRCV
fgDetail.Col = 8: fgDetail.Text = xYSWIDOS0.SWIDOS52A
fgDetail.Col = 9: fgDetail.Text = xYSWIDOS0.SWIDOS50PI
fgDetail.Col = 10: fgDetail.Text = xYSWIDOS0.SWIDOS57A
fgDetail.Col = 11: fgDetail.Text = xYSWIDOS0.SWIDOS59PI
fgDetail.Col = 12: fgDetail.Text = xYSWIDOS0.SWIDOS20
fgDetail.Col = 13: fgDetail.Text = xYSWIDOS0.SWIDOS21
fgDetail.Col = 14: fgDetail.Text = xYSWIDOS0.SWIDOSSABK


'fgDetail.Col = fgDetail_arrIndex: fgDetail.Text = lIndex
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



'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim wFct As String
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

wFct = UCase$(Trim(Mid$(Msg, 1, 12)))
Call BiaPgmAut_Init(wFct, YSWIDOS0_Aut)

'blnSetfocus = True
Form_Init


blnAuto = False


End Sub


Public Sub Form_Init()
Dim V, xSql As String, X As String
Dim K As Long

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True

blnControl = False

cmdReset


fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False


fraSelect_Options_1.BorderStyle = 0


lstW.Visible = False
Set lstW.Container = fraTab0
lstW.Top = fgDetail.Top + 300
lstW.Left = fgDetail.Left + fgDetail.Width - lstW.Width - 300
lstW.Height = fgDetail.Height - 300
lstW.BackColor = &HFAFAFA
lstW.ForeColor = vbBlack '&H4080&

fgDetail_FormatString = fgDetail.FormatString
wAmjMin = Mid$(YBIATAB0_DATE_CPT_AP1, 1, 4) & "0101"
wAmjMax = YBIATAB0_DATE_CPT_AP1
Call DTPicker_Set(txtSelect_SWIDOSAMJ_Max, wAmjMax) '
Call DTPicker_Set(txtSelect_SWIDOSAMJ_Min, wAmjMin) '

cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1  - sélection (filtre)"
cboSelect_SQL.ListIndex = 0


lstW.Clear

param_Init

Me.Enabled = True
End Sub

Private Sub cboSelect_SWIDOS50PI_K_Click()
If Trim(cboSelect_SWIDOS50PI_K) = "" Then cboSelect_SWIDOS50PI.ListIndex = 0

End Sub


Private Sub cboSelect_SWIDOS52A_K_Click()
If Trim(cboSelect_SWIDOS52A_K) = "" Then cboSelect_SWIDOS52A.ListIndex = 0

End Sub


Private Sub cboSelect_SWIDOS57A_K_Click()
If Trim(cboSelect_SWIDOS57A_K) = "" Then cboSelect_SWIDOS57A.ListIndex = 0

End Sub


Private Sub cboSelect_SWIDOS59PI_K_Click()
If Trim(cboSelect_SWIDOS59PI_K) = "" Then cboSelect_SWIDOS59PI.ListIndex = 0

End Sub


Private Sub cboSelect_SWIDOSBPIZ_K_Click()
If Trim(cboSelect_SWIDOSBPIZ_K) = "" Then cboSelect_SWIDOSBPIZ.ListIndex = 0

End Sub

Private Sub cboSelect_SWIDOSDEV_K_Click()
If Trim(cboSelect_SWIDOSDEV_K) = "" Then cboSelect_SWIDOSDEV.ListIndex = 0

End Sub


Private Sub cboSelect_SWIDOSDPIZ_K_Click()
If Trim(cboSelect_SWIDOSDPIZ_K) = "" Then cboSelect_SWIDOSDPIZ.ListIndex = 0

End Sub

Private Sub cboSelect_SWIDOSMTK_K_Click()
If Trim(cboSelect_SWIDOSMTK_K) = "" Then cboSelect_SWIDOSMTK.ListIndex = 0: cboSelect_SWIDOSMTK2.ListIndex = 0

End Sub


Private Sub cboSelect_SWIDOSOPEC_K_Click()
If Trim(cboSelect_SWIDOSOPEC_K) = "" Then cboSelect_SWIDOSOPEC.ListIndex = 0

End Sub

Private Sub cboSelect_SWIDOSRCV_K_Click()
If Trim(cboSelect_SWIDOSRCV_K) = "" Then cboSelect_SWIDOSRCV.ListIndex = 0

End Sub


Private Sub cboSelect_SWIDOSROUT_K_Click()
If Trim(cboSelect_SWIDOSROUT_K) = "" Then cboSelect_SWIDOSROUT.ListIndex = 0

End Sub


Private Sub cboSelect_SWIDOSSSE_K_Click()
If Trim(cboSelect_SWIDOSSSE_K) = "" Then cboSelect_SWIDOSSSE.ListIndex = 0

End Sub


Private Sub mnuExportation_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Exportation en cours ......"): DoEvents

cmdSelect_SQL_1

YSWIDOS0_Export


Me.Enabled = True: Me.MousePointer = 0
End Sub

Public Sub YSWIDOS0_Export()
On Error GoTo Error_Handler
Dim Nb As Long, wId As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSql As String
Dim wAmjMin As String, wAmjMax As String
Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim X As String, K As Long, K2 As Long, kMax As Long, K_Nb As Long, K_Mt As Long
Dim xWhere As String, X2 As String

Dim s_Solde As Currency, s_Prov As Currency
Dim t_Solde As Currency, t_Prov As Currency, x_Prov As Currency
Dim mSWIDOSDOS As Long, mSWIDOSAMJ As String
'______________________________________________
Call DTPicker_Control(txtSelect_SWIDOSAMJ_Min, wAmjMin)
Call DTPicker_Control(txtSelect_SWIDOSAMJ_Max, wAmjMax)

wFile = "C:\Temp\YSWIDOS0 " & DSys & " " & time_Hms & ".xlsx"
'______________________________________________

X = InputBox("par défaut : " & wFile _
    & vbCrLf & vbCrLf & "     =========================" _
    & vbCrLf & "     =========================", "SWI_STAT : nom du fichier d'exportation", wFile)
If Trim(X) = "" Then Exit Sub

wFilex = Trim(X)
'______________________________________________


If Dir(wFile) <> "" Then Kill wFile

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "SWI_STAT"
    .Subject = "SWI_STAT"
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "SWI_STAT"
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

Nb = 1
Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : " & Nb & " enregistrements"): DoEvents
For K = 1 To arrSWIDOS_Group_Nb
    wsExcel.Cells(Nb, K) = arrSWIDOS_Lib(arrSWIDOS_Group(K)): wsExcel.Columns(K).ColumnWidth = 9
    wsExcel.Columns(K).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
Next K

K_Nb = arrSWIDOS_Group_Nb + 1
wsExcel.Cells(Nb, K_Nb) = "Nombre": wsExcel.Columns(K_Nb).ColumnWidth = 10: wsExcel.Columns(K_Nb).NumberFormat = "#######"
wsExcel.Columns(K_Nb).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

K_Mt = arrSWIDOS_Group_Nb + 2
wsExcel.Cells(Nb, K_Mt) = "Montant"
wsExcel.Columns(K_Mt).ColumnWidth = 20: wsExcel.Columns(K_Mt).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(K_Mt).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight


For K = 1 To K_Mt
    wsExcel.Columns(K).VerticalAlignment = Excel.XlVAlign.xlVAlignTop
    wsExcel.Cells(1, K).Interior.Color = RGB(255, 255, 153)
Next K


For K = 1 To fgSelect.Rows - 1
    fgSelect.Row = K
    Nb = Nb + 1
    For K2 = 1 To arrSWIDOS_Group_Nb
        fgSelect.Col = K2 - 1
        wsExcel.Cells(Nb, K2) = fgSelect.Text
    Next K2
    fgSelect.Col = K_Nb - 1
    wsExcel.Cells(Nb, K_Nb) = Val(fgSelect.Text)
    fgSelect.Col = K_Mt - 1
    wsExcel.Cells(Nb, K_Mt) = CCur(num_CDec(fgSelect.Text))

Next K
    
Call lstErr_ChangeLastItem(lstErr, cmdContext, "Exportation en cours : " & Nb & " enregistrements"): DoEvents
Set rsSab = Nothing


wbExcel.SaveAs wFile

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

















'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------

blnControl = False
blnError = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
currentAction = ""
SSTab1.Tab = 0
blnControl = True

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
        For I = fgSelect_arrIndex To fgSelect.FixedCols Step -1
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
        For I = fgDetail_arrIndex To fgDetail.FixedCols Step -1
          fgDetail.Col = I: fgDetail.CellBackColor = lColor
        Next I
    End If
End If
fgDetail.LeftCol = fgDetail.FixedCols
fgDetail.Visible = True
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





Private Sub cboSelect_SWIDOSMTK_Click()
cmdSelect_Reset
End Sub

Private Sub cboSelect_SQL_Click()
cmdSelect_Reset

End Sub


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Dim X As String, I As Integer
Me.Enabled = False: Me.MousePointer = vbHourglass

Select Case SSTab1.Tab
    Case 0:
        Me.PopupMenu mnuPrint, vbPopupMenuLeftButton
    End Select

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdPrint_YSWIDOS0(blnDetail As Boolean)
Dim X As String, xSql As String, I As Integer, K As Integer
Dim wAmj As String, xWhere As String
Dim soldeD As typeYSWIDOS0, soldeF As typeYSWIDOS0, Total As typeYSWIDOS0
Dim blnXprt_Line As Boolean
Dim Nb_Detail As Long


End Sub

Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SWI_STAT_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Reset
fgSelect.Visible = False
fraSelect_Options.Visible = False

Select Case cmdSelect_SQL_K
    Case "1": fraSelect_Options.Visible = True: cmdSelect_SQL_1
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< SWI_STAT_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus

End Sub


Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wOrigine As String
On Error Resume Next


If Y <= fgDetail.RowHeightMin Then
Else
    If fgDetail.Rows > 1 Then
       ' blnControl = False
        Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
        fgDetail.Col = fgDetail_arrIndex:  arrYSWIDOS0_Index = CLng(fgDetail.Text)
        fraDetail_Display
   End If
End If

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
    Case Is = 27: cmdContext_Quit: KeyCode = 0
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select


End Sub

Public Sub cmdContext_Quit()
'blnControl = False
lstErr.Clear: lstErr.Height = 200

If SSTab1.Tab <> 0 Then
    SSTab1.Tab = 0
    Exit Sub
End If
If lstW.Visible Then
    lstW.Visible = False
    Exit Sub
End If

If fgDetail.Visible Then
    fgDetail.Visible = False
    Exit Sub
End If
If fgSelect.Visible Then
    fgSelect.Visible = False
    Exit Sub
End If

If SSTab1.Tab = 0 Then
    Unload Me
End If

End Sub
Public Sub cmdContext_Return()
    If SSTab1.Tab = 0 Then
        If cmdSelect_SQL_K <> "J" And cmdSelect_SQL_K <> "J#" Then
            If Not fgSelect.Version Then cmdSelect_Ok_Click
        End If
    Else
        SendKeys "{TAB}"
    End If
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





Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wOrigine As String, xSql As String
On Error Resume Next

fgDetail.Visible = False
lstW.Visible = False
If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  arrYSWIDOS0_Index = CLng(fgSelect.Text)
        
    fgDetail_Display
        
   End If
End If

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


Public Sub fgDetail_Reset()
fgDetail.Clear
fgDetail_Sort1 = 0: fgDetail_Sort2 = 0
fgDetail_Sort1_Old = -1
fgDetail_RowDisplay = 0: fgDetail_RowClick = 0
fgDetail_arrIndex = fgDetail.Cols - 1
blnfgDetail_DisplayLine = False
fgDetail_SortAD = 6
fgDetail.LeftCol = fgDetail.FixedCols

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






Public Sub fgSelect_ForeColor(lColor As Long)
For I = 0 To fgSelect_arrIndex
  fgSelect.Col = I: fgSelect.CellForeColor = lColor
Next I

End Sub
















Private Sub txtSelect_SWIDOS20_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtSElect_SWIDOS21_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtSelect_SWIDOSAMJ_Max_Change()
If fgSelect.Visible Then cmdSelect_Reset

End Sub

Private Sub txtSelect_SWIDOSAMJ_Max_KeyPress(KeyAscii As Integer)
cmdSelect_Reset

End Sub

Private Sub txtSelect_SWIDOSAMJ_Min_Change()
If fgSelect.Visible Then cmdSelect_Reset

End Sub


Private Sub txtSelect_SWIDOSAMJ_Min_KeyPress(KeyAscii As Integer)
cmdSelect_Reset

End Sub

Public Sub param_Init()
Dim xSql As String
Dim K As Integer

arrSWIDOS_Field(0) = "": arrSWIDOS_Lib(1) = ""
arrSWIDOS_Field(1) = "SWIDOSOPEC": arrSWIDOS_Lib(1) = "Code Opé"
arrSWIDOS_Field(2) = "SWIDOSMTK": arrSWIDOS_Lib(2) = "Type MT"
arrSWIDOS_Field(3) = "SWIDOSDEV": arrSWIDOS_Lib(3) = "Devise"
arrSWIDOS_Field(4) = "SWIDOSSSE": arrSWIDOS_Lib(4) = "Service"
arrSWIDOS_Field(5) = "SWIDOSRCV": arrSWIDOS_Lib(5) = "BIC Receveur"
arrSWIDOS_Field(6) = "SWIDOS52A": arrSWIDOS_Lib(6) = "BIC BQ  D.O."
arrSWIDOS_Field(7) = "SWIDOS57A": arrSWIDOS_Lib(7) = "BIC BQ  BEN"
arrSWIDOS_Field(8) = "SWIDOSROUT": arrSWIDOS_Lib(8) = "Route"
arrSWIDOS_Field(9) = "SWIDOS50PI": arrSWIDOS_Lib(9) = "Pays DO"
arrSWIDOS_Field(10) = "SWIDOS59PI": arrSWIDOS_Lib(10) = "Pays BEN"
arrSWIDOS_Field(11) = "SWIDOSDPIZ": arrSWIDOS_Lib(11) = "Zone Pays DO"
arrSWIDOS_Field(12) = "SWIDOSBPIZ": arrSWIDOS_Lib(12) = "Zone Pays BEN"
arrSWIDOS_Field_Nb = 12

For K = 0 To arrSWIDOS_Field_Nb
    If K > 0 Then cboSelect_SWIDOS_1.AddItem arrSWIDOS_Lib(K)
    cboSelect_SWIDOS_2.AddItem arrSWIDOS_Lib(K)
    cboSelect_SWIDOS_3.AddItem arrSWIDOS_Lib(K)
    cboSelect_SWIDOS_4.AddItem arrSWIDOS_Lib(K)
    cboSelect_SWIDOS_5.AddItem arrSWIDOS_Lib(K)
Next K
cboSelect_SWIDOS_1.ListIndex = 0
arrSWIDOS_Group(1) = 0
arrSWIDOS_Group_Nb = 0

'_______________________________________________________________________________
cboSelect_SWIDOSOPEC_K.Clear
cboSelect_SWIDOSOPEC_K.AddItem ""
cboSelect_SWIDOSOPEC_K.AddItem "="
cboSelect_SWIDOSOPEC_K.AddItem "<>"

cboSelect_SWIDOSOPEC.Clear
cboSelect_SWIDOSOPEC.AddItem ""
xSql = "select distinct SWIDOSOPEC from " & paramIBM_Library_SABSPE & ".YSWIDOS0 order by SWIDOSOPEC"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWIDOSOPEC.AddItem Trim(rsSab("SWIDOSOPEC"))
    rsSab.MoveNext
Loop

'_______________________________________________________________________________
cboSelect_SWIDOSMTK_K.Clear
cboSelect_SWIDOSMTK_K.AddItem ""
cboSelect_SWIDOSMTK_K.AddItem "="
cboSelect_SWIDOSMTK_K.AddItem "<>"

cboSelect_SWIDOSMTK.Clear
cboSelect_SWIDOSMTK.AddItem ""
cboSelect_SWIDOSMTK2.Clear
cboSelect_SWIDOSMTK2.AddItem ""
xSql = "select distinct SWIDOSMTK from " & paramIBM_Library_SABSPE & ".YSWIDOS0 order by SWIDOSMTK"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWIDOSMTK.AddItem Trim(rsSab("SWIDOSMTK"))
    cboSelect_SWIDOSMTK2.AddItem Trim(rsSab("SWIDOSMTK"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWIDOSDEV_K.Clear
cboSelect_SWIDOSDEV_K.AddItem ""
cboSelect_SWIDOSDEV_K.AddItem "="
cboSelect_SWIDOSDEV_K.AddItem "<>"

cboSelect_SWIDOSDEV.Clear
cboSelect_SWIDOSDEV.AddItem ""
xSql = "select distinct SWIDOSDEV from " & paramIBM_Library_SABSPE & ".YSWIDOS0 order by SWIDOSDEV"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWIDOSDEV.AddItem Trim(rsSab("SWIDOSDEV"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWIDOSSSE_K.Clear
cboSelect_SWIDOSSSE_K.AddItem ""
cboSelect_SWIDOSSSE_K.AddItem "="
cboSelect_SWIDOSSSE_K.AddItem "<>"

cboSelect_SWIDOSSSE.Clear
cboSelect_SWIDOSSSE.AddItem ""
xSql = "select distinct SWIDOSSSE from " & paramIBM_Library_SABSPE & ".YSWIDOS0 order by SWIDOSSSE"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWIDOSSSE.AddItem Trim(rsSab("SWIDOSSSE"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWIDOSRCV_K.Clear
cboSelect_SWIDOSRCV_K.AddItem ""
cboSelect_SWIDOSRCV_K.AddItem "="
cboSelect_SWIDOSRCV_K.AddItem "=4"
cboSelect_SWIDOSRCV_K.AddItem "=6"
cboSelect_SWIDOSRCV_K.AddItem "=8"
cboSelect_SWIDOSRCV_K.AddItem "<>"

cboSelect_SWIDOSRCV.Clear
cboSelect_SWIDOSRCV.AddItem ""
xSql = "select distinct SWIDOSRCV from " & paramIBM_Library_SABSPE & ".YSWIDOS0 order by SWIDOSRCV"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWIDOSRCV.AddItem Trim(rsSab("SWIDOSRCV"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWIDOS52A_K.Clear
cboSelect_SWIDOS52A_K.AddItem ""
cboSelect_SWIDOS52A_K.AddItem "="
cboSelect_SWIDOS52A_K.AddItem "=4"
cboSelect_SWIDOS52A_K.AddItem "=6"
cboSelect_SWIDOS52A_K.AddItem "=8"
cboSelect_SWIDOS52A_K.AddItem "<>"

cboSelect_SWIDOS52A.Clear
cboSelect_SWIDOS52A.AddItem ""
xSql = "select distinct SWIDOS52A from " & paramIBM_Library_SABSPE & ".YSWIDOS0 order by SWIDOS52A"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWIDOS52A.AddItem Trim(rsSab("SWIDOS52A"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWIDOS57A_K.Clear
cboSelect_SWIDOS57A_K.AddItem ""
cboSelect_SWIDOS57A_K.AddItem "="
cboSelect_SWIDOS57A_K.AddItem "=4"
cboSelect_SWIDOS57A_K.AddItem "=6"
cboSelect_SWIDOS57A_K.AddItem "=8"
cboSelect_SWIDOS57A_K.AddItem "<>"

cboSelect_SWIDOS57A.Clear
cboSelect_SWIDOS57A.AddItem ""
xSql = "select distinct SWIDOS57A from " & paramIBM_Library_SABSPE & ".YSWIDOS0 order by SWIDOS57A"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWIDOS57A.AddItem Trim(rsSab("SWIDOS57A"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWIDOSROUT_K.Clear
cboSelect_SWIDOSROUT_K.AddItem ""
cboSelect_SWIDOSROUT_K.AddItem "="
cboSelect_SWIDOSROUT_K.AddItem "<>"

cboSelect_SWIDOSROUT.Clear
cboSelect_SWIDOSROUT.AddItem ""
xSql = "select distinct SWIDOSROUT from " & paramIBM_Library_SABSPE & ".YSWIDOS0 order by SWIDOSROUT"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWIDOSROUT.AddItem Trim(rsSab("SWIDOSROUT"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWIDOS50PI_K.Clear
cboSelect_SWIDOS50PI_K.AddItem ""
cboSelect_SWIDOS50PI_K.AddItem "="
cboSelect_SWIDOS50PI_K.AddItem "<>"

cboSelect_SWIDOS50PI.Clear
cboSelect_SWIDOS50PI.AddItem ""
xSql = "select distinct SWIDOS50PI from " & paramIBM_Library_SABSPE & ".YSWIDOS0 order by SWIDOS50PI"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWIDOS50PI.AddItem Trim(rsSab("SWIDOS50PI"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWIDOS59PI_K.Clear
cboSelect_SWIDOS59PI_K.AddItem ""
cboSelect_SWIDOS59PI_K.AddItem "="
cboSelect_SWIDOS59PI_K.AddItem "<>"

cboSelect_SWIDOS59PI.Clear
cboSelect_SWIDOS59PI.AddItem ""
xSql = "select distinct SWIDOS59PI from " & paramIBM_Library_SABSPE & ".YSWIDOS0 order by SWIDOS59PI"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWIDOS59PI.AddItem Trim(rsSab("SWIDOS59PI"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWIDOSDPIZ_K.Clear
cboSelect_SWIDOSDPIZ_K.AddItem ""
cboSelect_SWIDOSDPIZ_K.AddItem "="
cboSelect_SWIDOSDPIZ_K.AddItem "<>"

cboSelect_SWIDOSDPIZ.Clear
cboSelect_SWIDOSDPIZ.AddItem ""
cboSelect_SWIDOSDPIZ.AddItem "FR"
cboSelect_SWIDOSDPIZ.AddItem "UE"
cboSelect_SWIDOSDPIZ.AddItem "**"
'_______________________________________________________________________________
cboSelect_SWIDOSBPIZ_K.Clear
cboSelect_SWIDOSBPIZ_K.AddItem ""
cboSelect_SWIDOSBPIZ_K.AddItem "="
cboSelect_SWIDOSBPIZ_K.AddItem "<>"

cboSelect_SWIDOSBPIZ.Clear
cboSelect_SWIDOSBPIZ.AddItem ""
cboSelect_SWIDOSBPIZ.AddItem "FR"
cboSelect_SWIDOSBPIZ.AddItem "UE"
cboSelect_SWIDOSBPIZ.AddItem "**"

End Sub

Public Sub fraDetail_Display()
Dim xSql As String, xWhere As String
On Error Resume Next
fgDetail.Col = 14
xWhere = " where SWIHIBETA  = " & currentZMNURUT0.MNURUTETB _
     & " and   SWIHIBNUM = " & fgDetail.Text _
     & " order by SWIHIBNEN , SWIHIBNLI"

xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIHIB0 " & xWhere
Set rsSab = cnsab.Execute(xSql)
lstW.Clear
Do While Not rsSab.EOF
    lstW.AddItem rsSab("SWIHIBDET")
    rsSab.MoveNext
Loop


lstW.Visible = True
lstW.SetFocus
End Sub
