VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8532
   ClientLeft      =   108
   ClientTop       =   408
   ClientWidth     =   10944
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8532
   ScaleWidth      =   10944
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   8052
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   10212
      _ExtentX        =   18013
      _ExtentY        =   14203
      _Version        =   393216
      TabOrientation  =   3
      TabHeight       =   420
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":0582
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fg"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Form1.frx":059E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgH"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgH 
         Height          =   4692
         Left            =   -73200
         TabIndex        =   2
         Top             =   1440
         Width           =   7812
         _ExtentX        =   13780
         _ExtentY        =   8276
         _Version        =   393216
         Rows            =   10
         Cols            =   3
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "@Arial Unicode MS"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin MSFlexGridLib.MSFlexGrid fg 
         Height          =   5652
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   9132
         _ExtentX        =   16108
         _ExtentY        =   9970
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         RowHeightMin    =   300
         BackColor       =   16777215
         ForeColor       =   12582912
         BackColorBkg    =   12632256
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         GridLineWidth   =   2
         FormatString    =   "<Echéance   |<Nature |Informations                                                                            "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "@Arial Unicode MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
fg.WordWrap = True
fg.Rows = 10
fg.Row = 1
fg.Col = 0: fg.Text = "30-06-2009": fg.CellForeColor = vbMagenta
fg.CellBackColor = &HC0E0FF
fg.Col = 1: fg.Text = "P 1"
fg.CellBackColor = &HC0E0FF
fg.Col = 2: fg.Text = "LOULERGUE le 18-08-2009 10:54"
fg.CellBackColor = &HC0E0FF
fg.CellFontSize = 7
fg.Row = 2: fg.Col = 2
fg.Text = "aaaaaaaaaaaaaaaaaaaa" & vbCrLf _
       & "zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz" & vbCrLf _
       & "eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee" & vbCrLf _
       & "rrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqqq"
fg.RowHeight(2) = 1200
       
fg.Row = 3
fg.Col = 0: fg.Text = "30-06-2009": fg.CellForeColor = &H4000&
fg.CellBackColor = &HC0FFFF
fg.Col = 1: fg.Text = "A 1.1"
fg.CellBackColor = &HC0FFFF
fg.Col = 2: fg.Text = "LOULERGUE le 18-08-2009 10:54"
fg.CellBackColor = &HC0FFFF
fg.CellFontSize = 7
fg.Row = 4: fg.Col = 2
fg.Text = "aaaaaaaaaaaaaaaaaaaa" & vbCrLf _
       & "zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz" & vbCrLf _
       & "eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee" & vbCrLf _
       & "rrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqqq"
fg.RowHeight(4) = 1200
       
End Sub


