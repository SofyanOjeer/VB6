VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmYTP7OPH0 
   AutoRedraw      =   -1  'True
   Caption         =   "Flux_NOSTRO"
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
   Icon            =   "YTP7OPH0.frx":0000
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Stat comptes Nostro, caisse "
      TabPicture(0)   =   "YTP7OPH0.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "."
      TabPicture(1)   =   "YTP7OPH0.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstW"
      Tab(1).ControlCount=   1
      Begin VB.ListBox lstW 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   -73920
         TabIndex        =   14
         Top             =   3552
         Visible         =   0   'False
         Width           =   4212
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
         Begin MSFlexGridLib.MSFlexGrid fgCPTPIE 
            Height          =   3180
            Left            =   2640
            TabIndex        =   26
            Top             =   6120
            Visible         =   0   'False
            Width           =   10332
            _ExtentX        =   18230
            _ExtentY        =   5609
            _Version        =   393216
            Cols            =   7
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   16777215
            ForeColor       =   4210752
            BackColorFixed  =   14737632
            ForeColorFixed  =   16384
            BackColorBkg    =   -2147483633
            AllowUserResizing=   3
            FormatString    =   $"YTP7OPH0.frx":0342
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
         Begin VB.Frame fraDetail 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   492
            Left            =   3960
            TabIndex        =   16
            Top             =   1440
            Visible         =   0   'False
            Width           =   8292
            Begin VB.Label libCOMPTEDEV 
               BackColor       =   &H00C0FFC0&
               Caption         =   "comptedev"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   120
               TabIndex        =   19
               Top             =   120
               Width           =   492
            End
            Begin VB.Label libTP7OPHCOM 
               BackColor       =   &H00C0FFC0&
               Caption         =   "libellé"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   252
               Left            =   3000
               TabIndex        =   18
               Top             =   120
               Width           =   3252
            End
            Begin VB.Label lblTP7OPHCOM 
               BackColor       =   &H00C0FFC0&
               Caption         =   "compte"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   1080
               TabIndex        =   17
               Top             =   120
               Width           =   1812
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
            Left            =   9360
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   360
            Width           =   3732
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
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   840
            Width           =   1335
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
            Height          =   1212
            Left            =   360
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   8712
            Begin VB.ComboBox cboSelect_TP7OPHOPE 
               Height          =   312
               Left            =   5280
               Sorted          =   -1  'True
               TabIndex        =   23
               Text            =   "OPE"
               Top             =   480
               Width           =   1176
            End
            Begin VB.Frame fraSelect_Options_1 
               BackColor       =   &H00F0FFFF&
               BorderStyle     =   0  'None
               Height          =   852
               Left            =   0
               TabIndex        =   11
               Top             =   120
               Width           =   4932
               Begin VB.ComboBox cboSelect_TP7OPHCOM 
                  Height          =   312
                  Left            =   120
                  Sorted          =   -1  'True
                  TabIndex        =   22
                  Text            =   "compte"
                  Top             =   360
                  Width           =   2136
               End
               Begin VB.ComboBox cboSelect_TP7OPHDEV 
                  Height          =   312
                  Left            =   3120
                  Sorted          =   -1  'True
                  TabIndex        =   20
                  Text            =   "dev"
                  Top             =   360
                  Width           =   1176
               End
               Begin VB.Label lblSelect_TP7OPHDEV 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "devise"
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
                  Left            =   3240
                  TabIndex        =   21
                  Top             =   0
                  Width           =   612
               End
               Begin VB.Label lblSelect_TP7OPHCOM 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "compte"
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
                  Left            =   840
                  TabIndex        =   12
                  Top             =   0
                  Width           =   612
               End
            End
            Begin MSComCtl2.DTPicker txtSelect_TP7OPHDTR_Max 
               Height          =   300
               Left            =   7320
               TabIndex        =   9
               Top             =   840
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
               Format          =   104726531
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_TP7OPHDTR_Min 
               Height          =   300
               Left            =   7320
               TabIndex        =   15
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
               Format          =   104726531
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_TP7OPHOPE 
               BackColor       =   &H00F0FFFF&
               Caption         =   "opération"
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
               Left            =   5520
               TabIndex        =   24
               Top             =   120
               Width           =   972
            End
            Begin VB.Label lblSelect_TP7OPHDTR 
               BackColor       =   &H00F0FFFF&
               Caption         =   "date de situation"
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
               Left            =   7320
               TabIndex        =   10
               Top             =   120
               Width           =   1212
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   4548
            Left            =   360
            TabIndex        =   5
            Top             =   1440
            Visible         =   0   'False
            Width           =   2712
            _ExtentX        =   4789
            _ExtentY        =   8017
            _Version        =   393216
            Rows            =   1
            Cols            =   4
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   -2147483633
            ForeColor       =   16384
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483637
            BackColorSel    =   12648384
            BackColorBkg    =   -2147483633
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   "<Dev  |<Compte                          ||"
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
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   4020
            Left            =   3960
            TabIndex        =   13
            Top             =   1920
            Visible         =   0   'False
            Width           =   8412
            _ExtentX        =   14843
            _ExtentY        =   7091
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   16777215
            ForeColor       =   16384
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483633
            BackColorBkg    =   -2147483633
            FormatString    =   $"YTP7OPH0.frx":040C
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
         Begin MSFlexGridLib.MSFlexGrid fgBIAMVT 
            Height          =   3300
            Left            =   360
            TabIndex        =   25
            Top             =   6000
            Visible         =   0   'False
            Width           =   12612
            _ExtentX        =   22251
            _ExtentY        =   5821
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   16777215
            ForeColor       =   16384
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483633
            BackColorBkg    =   -2147483633
            AllowUserResizing=   3
            FormatString    =   $"YTP7OPH0.frx":049A
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
      Picture         =   "YTP7OPH0.frx":058F
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
      Begin VB.Menu mnuPrint_Recap 
         Caption         =   "Etat récapitulatif"
      End
      Begin VB.Menu mnuPrint_Detail 
         Caption         =   "Etat détaillé"
      End
   End
End
Attribute VB_Name = "frmYTP7OPH0"
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
Dim YTP7OPH0_Aut As typeAuthorization
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
Dim xYBIACPT0 As typeYBIACPT0, newYBIACPT0 As typeYBIACPT0, oldYBIACPT0 As typeYBIACPT0
Dim arrYBIACPT0() As typeYBIACPT0, arrYBIACPT0_Nb As Long, arrYBIACPT0_Max As Long, arrYBIACPT0_Index As Long

Dim xYTP7OPH0 As typeYTP7OPH0, newYTP7OPH0 As typeYTP7OPH0, oldYTP7OPH0 As typeYTP7OPH0
Dim arrYTP7OPH0() As typeYTP7OPH0, arrYTP7OPH0_Nb As Long, arrYTP7OPH0_Max As Long, arrYTP7OPH0_Index As Long

Dim fgDetail_FormatString As String, fgDetail_K As Integer
Dim fgDetail_RowDisplay As Integer, fgDetail_RowClick As Integer, fgDetail_ColClick As Integer
Dim fgDetail_ColorClick As Long, fgDetail_ColorDisplay As Long
Dim fgDetail_Sort1 As Integer, fgDetail_Sort2 As Integer
Dim fgDetail_SortAD As Integer, fgDetail_Sort1_Old As Integer
Dim fgDetail_arrIndex As Integer
Dim blnfgDetail_DisplayLine As Boolean



Dim fgBIAMVT_FormatString As String, fgBIAMVT_K As Integer
Dim fgBIAMVT_RowDisplay As Integer, fgBIAMVT_RowClick As Integer, fgBIAMVT_ColClick As Integer
Dim fgBIAMVT_ColorClick As Long, fgBIAMVT_ColorDisplay As Long
Dim fgBIAMVT_Sort1 As Integer, fgBIAMVT_Sort2 As Integer
Dim fgBIAMVT_SortAD As Integer, fgBIAMVT_Sort1_Old As Integer
Dim fgBIAMVT_arrIndex As Integer
Dim blnfgBIAMVT_DisplayLine As Boolean

Dim xYBIAMVTH As typeYBIAMVT0, newYBIAMVTH As typeYBIAMVT0, oldYBIAMVTH As typeYBIAMVT0
Dim arrYBIAMVTH() As typeYBIAMVT0, arrYBIAMVTH_Nb As Long, arrYBIAMVTH_Max As Long, arrYBIAMVTH_Index As Long

Dim fgCPTPIE_FormatString As String, fgCPTPIE_K As Integer
Dim fgCPTPIE_RowDisplay As Integer, fgCPTPIE_RowClick As Integer, fgCPTPIE_ColClick As Integer
Dim fgCPTPIE_ColorClick As Long, fgCPTPIE_ColorDisplay As Long
Dim fgCPTPIE_Sort1 As Integer, fgCPTPIE_Sort2 As Integer
Dim fgCPTPIE_SortAD As Integer, fgCPTPIE_Sort1_Old As Integer
Dim fgCPTPIE_arrIndex As Integer
Dim blnfgCPTPIE_DisplayLine As Boolean

Dim xYCPTPIEH As typeYBIAMVT0, newYCPTPIEH As typeYBIAMVT0, oldYCPTPIEH As typeYBIAMVT0
Dim arrYCPTPIEH() As typeYBIAMVT0, arrYCPTPIEH_Nb As Long, arrYCPTPIEH_Max As Long, arrYCPTPIEH_Index As Long

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim arrDev() As String, arrDev_Nb As Integer
Dim arrOPE() As String, arrOPE_Nb As Integer, arrOPE_K As Integer


'______________________________________________________________________
Private Sub fgSelect_Display()
Dim wColor As Long

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
    
For I = 1 To arrYBIACPT0_Nb
         
    xYBIACPT0 = arrYBIACPT0(I)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine I
    
Next I

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYBIACPT0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub arrYTP7OPH0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYTP7OPH0(101)
arrYTP7OPH0_Max = 100: arrYTP7OPH0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YTP7OPH0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYTP7OPH0_GetBuffer(rsSab, xYTP7OPH0)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYTP7OPH0.fgselect_Display"
        '' Exit Sub
     Else
         arrYTP7OPH0_Nb = arrYTP7OPH0_Nb + 1
         If arrYTP7OPH0_Nb > arrYTP7OPH0_Max Then
             arrYTP7OPH0_Max = arrYTP7OPH0_Max + 100
             ReDim Preserve arrYTP7OPH0(arrYTP7OPH0_Max)
         End If
         
         arrYTP7OPH0(arrYTP7OPH0_Nb) = xYTP7OPH0
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

Private Sub arrYBIACPT0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYBIACPT0(101)
arrYBIACPT0_Max = 100: arrYBIACPT0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYBIACPT0_GetBuffer(rsSab, xYBIACPT0)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYBIACPT0.fgselect_Display"
        '' Exit Sub
     Else
         arrYBIACPT0_Nb = arrYBIACPT0_Nb + 1
         If arrYBIACPT0_Nb > arrYBIACPT0_Max Then
             arrYBIACPT0_Max = arrYBIACPT0_Max + 100
             ReDim Preserve arrYBIACPT0(arrYBIACPT0_Max)
         End If
         
         arrYBIACPT0(arrYBIACPT0_Nb) = xYBIACPT0
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


Private Sub arrYBIAMVTH_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYBIAMVTH(101)
arrYBIAMVTH_Max = 100: arrYBIAMVTH_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVTH)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYBIAMVTH.fgselect_Display"
        '' Exit Sub
     Else
         arrYBIAMVTH_Nb = arrYBIAMVTH_Nb + 1
         If arrYBIAMVTH_Nb > arrYBIAMVTH_Max Then
             arrYBIAMVTH_Max = arrYBIAMVTH_Max + 100
             ReDim Preserve arrYBIAMVTH(arrYBIAMVTH_Max)
         End If
         
         arrYBIAMVTH(arrYBIAMVTH_Nb) = xYBIAMVTH
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
    fgDetail.Visible = False: fraDetail.Visible = False
    fgBIAMVT.Visible = False
    lstW.Visible = False
    cmdSelect_Ok.Visible = False 'True
    cmdSelect_SQL_K = Trim(mId$(cboSelect_SQL, 1, 2))
    Select Case cmdSelect_SQL_K
        Case "1":
            fraSelect_Options.Visible = True: fraSelect_Options_1.Visible = True
            cmdSelect_Ok_Click
        Case Else
            cmdSelect_Ok.Visible = True
    End Select

End If

End Sub

Public Sub cmdDetail_Reset()
If blnControl Then
    lstErr.Clear
    If fgDetail.Visible Then
        fgDetail.Visible = False: fraDetail.Visible = False
        fgBIAMVT.Visible = False
        fgCPTPIE.Visible = False
        fgDetail_Display
    End If
End If

End Sub


Private Sub cmdSelect_SQL_1()
Dim V
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdYTP7OPH0_SQL"
blnOk = False

ReDim arrYBIACPT0(cboSelect_TP7OPHCOM.ListCount)
arrYBIACPT0_Nb = 0

xWhere = ""
X = Trim(cboSelect_TP7OPHCOM)
If X <> "" Then xWhere = "   and TP7OPHCOM like '%" & X & "%'"
X = Trim(cboSelect_TP7OPHDEV)
If X <> "" Then xWhere = xWhere & "   and TP7OPHDEV = '" & X & "'"
If xWhere <> "" Then Mid$(xWhere, 1, 6) = " where"

xSql = "select distinct TP7OPHDEV , TP7OPHCOM from " & paramIBM_Library_SABSPE & ".YTP7OPH0 " _
     & xWhere & " order by TP7OPHDEV , TP7OPHCOM"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    arrYBIACPT0_Nb = arrYBIACPT0_Nb + 1
    arrYBIACPT0(arrYBIACPT0_Nb).COMPTECOM = Trim(rsSab("TP7OPHCOM"))
    arrYBIACPT0(arrYBIACPT0_Nb).COMPTEDEV = Trim(rsSab("TP7OPHdev"))
    rsSab.MoveNext
Loop


'arrYBIACPT0_SQL xWhere & " order by COMPTEDEV , COMPTECOM "

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
Dim X As String, xWhere As String, xOPE As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAMJ As String

On Error GoTo Error_Handler
fgDetail.Visible = False: fraDetail.Visible = False
fgBIAMVT.Visible = False
fgCPTPIE.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString
fgDetail.Row = 0

currentAction = "fgDetail_Display"
Call DTPicker_Control(txtSelect_TP7OPHDTR_Min, wAmjMin)
Call DTPicker_Control(txtSelect_TP7OPHDTR_Max, wAmjMax)

lblTP7OPHCOM = oldYBIACPT0.COMPTECOM
xWhere = "select COMPTEINT , COMPTEDEV from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " Where COMPTECOM = '" & oldYBIACPT0.COMPTECOM & "'"
Set rsSab = cnsab.Execute(xWhere)
If Not rsSab.EOF Then
    libTP7OPHCOM = rsSab("COMPTEINT")
    libCOMPTEDEV = rsSab("COMPTEDEV")
Else
    libTP7OPHCOM = "??????"
End If
xOPE = ""
X = Trim(cboSelect_TP7OPHOPE)
If X <> "" Then xOPE = " and TP7OPHOPE = '" & X & "'"

xWhere = " where TP7OPHCOM ='" & Trim(oldYBIACPT0.COMPTECOM) & "'" _
     & xOPE _
     & " and TP7OPHDTR >= " & wAmjMin & " and TP7OPHDTR <= " & wAmjMax _
     & " order by TP7OPHDTR , TP7OPHOPE "

Call arrYTP7OPH0_SQL(xWhere)


For I = 1 To arrYTP7OPH0_Nb
         
    xYTP7OPH0 = arrYTP7OPH0(I)
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_DisplayLine I
    
Next I

fgDetail.Visible = True: fraDetail.Visible = True


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub fgBIAMVT_Display()
Dim wColor As Long
Dim X As String, xWhere As String, xOPE As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAMJ As String

On Error GoTo Error_Handler
fgBIAMVT.Visible = False
fgCPTPIE.Visible = False
fgBIAMVT_Reset

fgBIAMVT.Rows = 1
fgBIAMVT.FormatString = fgBIAMVT_FormatString
fgBIAMVT.Row = 0

currentAction = "fgBIAMVT_Display"

xWhere = " where MOUVEMCOM ='" & oldYTP7OPH0.TP7OPHCOM & "'" _
     & " and MOUVEMDTR = " & oldYTP7OPH0.TP7OPHDTR - 19000000 _
     & " and MOUVEMOPE = '" & oldYTP7OPH0.TP7OPHOPE & "'" _
     & " order by MOUVEMNUM , MOUVEMPIE , MOUVEMECR "

Call arrYBIAMVTH_SQL(xWhere)


For I = 1 To arrYBIAMVTH_Nb
         
    xYBIAMVTH = arrYBIAMVTH(I)
    fgBIAMVT.Rows = fgBIAMVT.Rows + 1
    fgBIAMVT.Row = fgBIAMVT.Rows - 1
    fgBIAMVT_DisplayLine I
    
Next I

fgBIAMVT.Visible = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgCPTPIE_Display()
Dim wColor As Long
Dim xSql As String, xWhere As String, xOPE As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAMJ As String

On Error GoTo Error_Handler
fgCPTPIE.Visible = False: fraDetail.Visible = False
fgCPTPIE_Reset

fgCPTPIE.Rows = 1
fgCPTPIE.FormatString = fgCPTPIE_FormatString
fgCPTPIE.Row = 0

currentAction = "fgCPTPIE_Display"

xWhere = " where MOUVEMETA =" & oldYBIAMVTH.MOUVEMETA _
     & " and MOUVEMPIE = " & oldYBIAMVTH.MOUVEMPIE _
     & " order by MOUVEMECR "


xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYBIAMVT0_GetBuffer(rsSab, xYCPTPIEH)
         
    fgCPTPIE.Rows = fgCPTPIE.Rows + 1
    fgCPTPIE.Row = fgCPTPIE.Rows - 1
    fgCPTPIE_DisplayLine I
    
    rsSab.MoveNext
Loop

fgCPTPIE.Visible = True: fraDetail.Visible = True


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub



Public Sub fgSelect_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim xSql As String
On Error Resume Next


fgSelect.Col = 0: fgSelect.Text = xYBIACPT0.COMPTEDEV
fgSelect.Col = 1: fgSelect.Text = xYBIACPT0.COMPTECOM

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
End Sub
Public Sub fgDetail_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim blnSolde As Boolean

On Error Resume Next
fgDetail.Col = 0: fgDetail.Text = dateImp10(xYTP7OPH0.TP7OPHDTR)
fgDetail.Col = 1: fgDetail.Text = xYTP7OPH0.TP7OPHOPE
fgDetail.Col = 2: fgDetail.Text = Format$(xYTP7OPH0.TP7OPHDBD, "### ### ### ##0.00")
fgBIAMVT.CellForeColor = vbRed
fgDetail.Col = 3: fgDetail.Text = Format$(xYTP7OPH0.TP7OPHDBN, "### ### ##0")
fgDetail.Col = 4: fgDetail.Text = Format$(xYTP7OPH0.TP7OPHCRD, "### ### ### ##0.00")
fgBIAMVT.CellForeColor = vbBlue
fgDetail.Col = 5: fgDetail.Text = Format$(xYTP7OPH0.TP7OPHCRN, "### ### ##0")

fgDetail.Col = fgDetail_arrIndex: fgDetail.Text = lIndex
End Sub


Public Sub fgBIAMVT_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim blnSolde As Boolean

On Error Resume Next
fgBIAMVT.Col = 0: fgBIAMVT.Text = xYBIAMVTH.MOUVEMSER & " " & xYBIAMVTH.MOUVEMSSE & " " & xYBIAMVTH.MOUVEMOPE & " " & xYBIAMVTH.MOUVEMEVE & " " & xYBIAMVTH.MOUVEMNUM
'fgBIAMVT.ForeColor = fgBIAMVT.BackColorFixed
fgBIAMVT.Col = 1: fgBIAMVT.Text = xYBIAMVTH.MOUVEMCOM
If xYBIAMVTH.MOUVEMMON > 0 Then
    fgBIAMVT.Col = 2: fgBIAMVT.Text = Format$(xYBIAMVTH.MOUVEMMON, "### ### ### ##0.00")
    fgBIAMVT.CellForeColor = vbRed
Else
    fgBIAMVT.Col = 3: fgBIAMVT.Text = Format$(Abs(xYBIAMVTH.MOUVEMMON), "### ### ### ##0.00")
    fgBIAMVT.CellForeColor = vbBlue
End If

fgBIAMVT.Col = 4: fgBIAMVT.Text = xYBIAMVTH.LIBELLIB1 & xYBIAMVTH.LIBELLIB2 & xYBIAMVTH.LIBELLIB3 & xYBIAMVTH.LIBELLIB4
fgBIAMVT.Col = 5: fgBIAMVT.Text = Format$(xYBIAMVTH.MOUVEMPIE, "### ### ##0") & "- " & Format$(xYBIAMVTH.MOUVEMECR, "### ##0")

fgBIAMVT.Col = fgBIAMVT_arrIndex: fgBIAMVT.Text = lIndex
End Sub


Public Sub fgCPTPIE_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim blnSolde As Boolean

On Error Resume Next
fgCPTPIE.Col = 0: fgCPTPIE.Text = xYCPTPIEH.MOUVEMCOM
If xYCPTPIEH.MOUVEMMON > 0 Then
    fgCPTPIE.Col = 1: fgCPTPIE.Text = Format$(xYCPTPIEH.MOUVEMMON, "### ### ### ##0.00")
    fgCPTPIE.CellForeColor = vbRed
Else
    fgCPTPIE.Col = 2: fgCPTPIE.Text = Format$(Abs(xYCPTPIEH.MOUVEMMON), "### ### ### ##0.00")
    fgCPTPIE.CellForeColor = vbBlue
End If

fgCPTPIE.Col = 3: fgCPTPIE.Text = xYCPTPIEH.LIBELLIB1 & xYCPTPIEH.LIBELLIB2 & xYCPTPIEH.LIBELLIB3 & xYCPTPIEH.LIBELLIB4
fgCPTPIE.Col = 4: fgCPTPIE.Text = Format$(xYCPTPIEH.MOUVEMECR, "### ##0")

fgCPTPIE.Col = fgCPTPIE_arrIndex: fgCPTPIE.Text = lIndex
End Sub



Public Sub YTP7OPH0_Export()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSql As String
Dim wAmjMin As String, wAmjMax As String
Dim X As String, K As Long, kMax As Long, K1 As Long, K2 As Long, K3 As Long
Dim xOPE As String, xOPE_where As String
Dim xDEV As String, xDEV_where As String
Dim Xcom As String, xCOM_where As String
Dim xDTR_where As String
'______________________________________________
xOPE_where = ""
xOPE = Trim(cboSelect_TP7OPHOPE)
If xOPE <> "" Then xOPE_where = " and TP7OPHOPE = '" & xOPE & "'"

xDEV_where = ""
xDEV = Trim(cboSelect_TP7OPHDEV)
If xDEV <> "" Then xDEV_where = " where TP7OPHDEV = '" & xDEV & "'"

xCOM_where = ""
Xcom = Trim(cboSelect_TP7OPHCOM)

xDTR_where = ""
If cmdSelect_SQL_K = "Ep" Then
    Call DTPicker_Control(txtSelect_TP7OPHDTR_Min, wAmjMin)
    Call DTPicker_Control(txtSelect_TP7OPHDTR_Max, wAmjMax)
    xDTR_where = " and TP7OPHDTR >= " & wAmjMin & " And TP7OPHDTR <= " & wAmjMax
End If

If Xcom = "" Then
    Select Case cmdSelect_SQL_K
        Case "En": Xcom = "N"
        Case "Ek": Xcom = "101100"
        Case "Ep": Xcom = "N"
    End Select
End If
xCOM_where = " and TP7OPHCOM like '" & Xcom & "%'"

wFile = Trim("C:\Temp\TP7 Stat " & xDEV & Xcom & xOPE & ".xlsx")
'______________________________________________

X = InputBox("par défaut : " & wFile _
    & vbCrLf & vbCrLf & "     =========================" _
    & vbCrLf & "     =========================", "Trésorerie prévisionnelle : nom du fichier d'exportation", wFile)
If Trim(X) = "" Then Exit Sub

wFilex = Trim(X)
'______________________________________________
If wFile <> wFilex Then
    wFile = wFilex
End If
'_________________________________________


If Dir(wFile) <> "" Then Kill wFile

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "Nostro"
    .Subject = ""
End With

'__________________________________________________________________________________
If xDEV <> "" Then

    Set wsExcel = wbExcel.ActiveSheet
    wsExcel.Name = xDEV
    
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YTP7OPH0 " _
         & " where TP7OPHDEV = '" & xDEV & "'" _
         & xOPE_where & xCOM_where & xDTR_where _
         & " order by TP7OPHDTR , TP7OPHOPE"
    Set rsSab = cnsab.Execute(xSql)
    If cmdSelect_SQL_K = "Ep" Then
        Call YTP7OPH0_Export_Feuille_DATE_OPE
    Else
        Call YTP7OPH0_Export_Feuille_Volume_Annuel
    End If
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "Exportation en cours : " & wsExcel.Name): DoEvents
    GoTo Exit_Sub
End If

'__________________________________________________________________________________
appExcel.Worksheets.Add , , arrDev_Nb - 3
For K = 1 To arrDev_Nb
    xDEV = arrDev(K)
    
    Set wsExcel = wbExcel.Sheets(K)
    wsExcel.Name = xDEV
    
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YTP7OPH0 " _
         & " where TP7OPHDEV = '" & xDEV & "'" _
         & xOPE_where & xCOM_where & xDTR_where _
         & " order by TP7OPHDTR , TP7OPHOPE"
    Set rsSab = cnsab.Execute(xSql)
    
    If cmdSelect_SQL_K = "Ep" Then
        Call YTP7OPH0_Export_Feuille_DATE_OPE
    Else
        Call YTP7OPH0_Export_Feuille_Volume_Annuel
    End If
    
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "Exportation en cours : " & wsExcel.Name): DoEvents

Next K
'__________________________________________________________________________________
Exit_Sub:
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


'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub


Public Sub YTP7OPH0_Export_Feuille_Volume_Annuel()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSql As String
Dim wAmjMin As String, wAmjMax As String
Dim X As String, K As Long, kMax As Long, K1 As Long, K2 As Long, K3 As Long
Dim wAAAA As Integer, wMM As Integer, wJJ As Integer

Dim TCR_Moy(20, 20) As Currency, TCR_Var(20, 20) As Double, TCR_Nb(20, 20) As Long
Dim TDB_Moy(20, 20) As Currency, TDB_Var(20, 20) As Double, TDB_Nb(20, 20) As Long
'______________________________________________

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

Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents

wsExcel.Cells(1, 1) = "Moyenne >": wsExcel.Columns(1).ColumnWidth = 10
wsExcel.Cells(1, 2) = "Année": wsExcel.Columns(2).ColumnWidth = 15: wsExcel.Columns(2).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 3) = "Janvier": wsExcel.Columns(3).ColumnWidth = 15: wsExcel.Columns(3).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 4) = "Février": wsExcel.Columns(4).ColumnWidth = 15: wsExcel.Columns(4).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 5) = "Mars": wsExcel.Columns(5).ColumnWidth = 15: wsExcel.Columns(5).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 6) = "Avril": wsExcel.Columns(6).ColumnWidth = 15: wsExcel.Columns(6).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 7) = "Mai": wsExcel.Columns(7).ColumnWidth = 15: wsExcel.Columns(7).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 8) = "Juin": wsExcel.Columns(8).ColumnWidth = 15: wsExcel.Columns(8).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 9) = "Juillet": wsExcel.Columns(9).ColumnWidth = 15: wsExcel.Columns(9).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 10) = "Août": wsExcel.Columns(10).ColumnWidth = 15: wsExcel.Columns(10).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 11) = "Septembre": wsExcel.Columns(11).ColumnWidth = 15: wsExcel.Columns(11).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 12) = "Octobre": wsExcel.Columns(12).ColumnWidth = 15: wsExcel.Columns(12).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 13) = "Novembre": wsExcel.Columns(13).ColumnWidth = 15: wsExcel.Columns(13).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"

wsExcel.Cells(1, 14) = "Décembre": wsExcel.Columns(14).ColumnWidth = 15: wsExcel.Columns(14).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 15) = "": wsExcel.Columns(15).ColumnWidth = 15: wsExcel.Columns(15).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 16) = "Lundi": wsExcel.Columns(16).ColumnWidth = 15: wsExcel.Columns(16).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 17) = "Mardi": wsExcel.Columns(17).ColumnWidth = 15: wsExcel.Columns(17).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 18) = "Mercredi": wsExcel.Columns(18).ColumnWidth = 15: wsExcel.Columns(18).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 19) = "Jeudi": wsExcel.Columns(19).ColumnWidth = 15: wsExcel.Columns(19).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Cells(1, 20) = "Vendredi": wsExcel.Columns(20).ColumnWidth = 15: wsExcel.Columns(20).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"

wsExcel.Cells(2, 1) = "2006 DB": wsExcel.Cells(2, 1).Font.Color = vbRed: wsExcel.Cells(3, 1) = "2006 CR"
wsExcel.Cells(4, 1) = "2007 DB": wsExcel.Cells(4, 1).Font.Color = vbRed: wsExcel.Cells(5, 1) = "2007 CR"
wsExcel.Cells(6, 1) = "2008 DB": wsExcel.Cells(6, 1).Font.Color = vbRed: wsExcel.Cells(7, 1) = "2008 CR"
wsExcel.Cells(8, 1) = "2009 DB": wsExcel.Cells(8, 1).Font.Color = vbRed: wsExcel.Cells(9, 1) = "2009 CR"
wsExcel.Cells(10, 1) = "2010 DB": wsExcel.Cells(10, 1).Font.Color = vbRed: wsExcel.Cells(11, 1) = "2010 CR"

For K = 1 To 20
    wsExcel.Cells(1, K).Interior.Color = RGB(0, 200, 255)
    wsExcel.Cells(13, K) = wsExcel.Cells(1, K)
    wsExcel.Cells(13, K).Interior.Color = RGB(0, 220, 255)
    wsExcel.Cells(25, K) = wsExcel.Cells(1, K)
    wsExcel.Cells(25, K).Interior.Color = RGB(0, 240, 255)
Next K
wsExcel.Cells(13, 1) = "Cumul >"
wsExcel.Cells(14, 1) = "2006 DB": wsExcel.Cells(15, 1) = "2006 CR"
wsExcel.Cells(14, 1).Font.Color = vbRed
wsExcel.Cells(16, 1) = "2007 DB": wsExcel.Cells(17, 1) = "2007 CR"
wsExcel.Cells(16, 1).Font.Color = vbRed
wsExcel.Cells(18, 1) = "2008 DB": wsExcel.Cells(19, 1) = "2008 CR"
wsExcel.Cells(18, 1).Font.Color = vbRed
wsExcel.Cells(20, 1) = "2009 DB": wsExcel.Cells(21, 1) = "2009 CR"
wsExcel.Cells(20, 1).Font.Color = vbRed
wsExcel.Cells(22, 1) = "2010 DB": wsExcel.Cells(23, 1) = "2010 CR"
wsExcel.Cells(22, 1).Font.Color = vbRed

wsExcel.Cells(25, 1) = "Nb >"
wsExcel.Cells(26, 1) = "2006 DB": wsExcel.Cells(27, 1) = "2006 CR"
wsExcel.Cells(26, 1).Font.Color = vbRed
wsExcel.Cells(28, 1) = "2007 DB": wsExcel.Cells(29, 1) = "2007 CR"
wsExcel.Cells(28, 1).Font.Color = vbRed
wsExcel.Cells(30, 1) = "2008 DB": wsExcel.Cells(31, 1) = "2008 CR"
wsExcel.Cells(30, 1).Font.Color = vbRed
wsExcel.Cells(32, 1) = "2009 DB": wsExcel.Cells(33, 1) = "2009 CR"
wsExcel.Cells(32, 1).Font.Color = vbRed
wsExcel.Cells(34, 1) = "2010 DB": wsExcel.Cells(35, 1) = "2010 CR"
wsExcel.Cells(34, 1).Font.Color = vbRed

rsYTP7OPH0_Init oldYTP7OPH0

For K1 = 0 To 20
    For K2 = 0 To 20
        TCR_Moy(K1, K2) = 0: TCR_Var(K1, K2) = 0: TCR_Nb(K1, K2) = 0
        TDB_Moy(K1, K2) = 0: TDB_Var(K1, K2) = 0: TDB_Nb(K1, K2) = 0
    Next K2
Next K1

'_______________________________________________________________________________________

Do While Not rsSab.EOF
    V = rsYTP7OPH0_GetBuffer(rsSab, xYTP7OPH0)
    If oldYTP7OPH0.TP7OPHDTR = xYTP7OPH0.TP7OPHDTR Then
        oldYTP7OPH0.TP7OPHDBN = oldYTP7OPH0.TP7OPHDBN + xYTP7OPH0.TP7OPHDBN
        oldYTP7OPH0.TP7OPHDBD = oldYTP7OPH0.TP7OPHDBD + xYTP7OPH0.TP7OPHDBD
        oldYTP7OPH0.TP7OPHCRN = oldYTP7OPH0.TP7OPHCRN + xYTP7OPH0.TP7OPHCRN
        oldYTP7OPH0.TP7OPHCRD = oldYTP7OPH0.TP7OPHCRD + xYTP7OPH0.TP7OPHCRD
    Else
        If oldYTP7OPH0.TP7OPHDTR > 0 Then
            wAAAA = Val(mId$(oldYTP7OPH0.TP7OPHDTR, 1, 4))
            wRow = (wAAAA - 2006)
            wMM = Val(mId$(oldYTP7OPH0.TP7OPHDTR, 5, 2))
            wJJ = 12 + Weekday(DateSerial(wAAAA, mId$(oldYTP7OPH0.TP7OPHDTR, 5, 2), mId$(oldYTP7OPH0.TP7OPHDTR, 7, 2)))
            TCR_Moy(wRow, 0) = TCR_Moy(wRow, 0) + oldYTP7OPH0.TP7OPHCRD
            TCR_Nb(wRow, 0) = TCR_Nb(wRow, 0) + oldYTP7OPH0.TP7OPHCRN
            TCR_Moy(wRow, wMM) = TCR_Moy(wRow, wMM) + oldYTP7OPH0.TP7OPHCRD
            TCR_Nb(wRow, wMM) = TCR_Nb(wRow, wMM) + oldYTP7OPH0.TP7OPHCRN
            TCR_Moy(wRow, wJJ) = TCR_Moy(wRow, wJJ) + oldYTP7OPH0.TP7OPHCRD
            TCR_Nb(wRow, wJJ) = TCR_Nb(wRow, wJJ) + oldYTP7OPH0.TP7OPHCRN
            
            TDB_Moy(wRow, 0) = TDB_Moy(wRow, 0) + oldYTP7OPH0.TP7OPHDBD
            TDB_Nb(wRow, 0) = TDB_Nb(wRow, 0) + oldYTP7OPH0.TP7OPHDBN
            TDB_Moy(wRow, wMM) = TDB_Moy(wRow, wMM) + oldYTP7OPH0.TP7OPHDBD
            TDB_Nb(wRow, wMM) = TDB_Nb(wRow, wMM) + oldYTP7OPH0.TP7OPHDBN
            TDB_Moy(wRow, wJJ) = TDB_Moy(wRow, wJJ) + oldYTP7OPH0.TP7OPHDBD
            TDB_Nb(wRow, wJJ) = TDB_Nb(wRow, wJJ) + oldYTP7OPH0.TP7OPHDBN
        End If
        oldYTP7OPH0 = xYTP7OPH0
    End If
    rsSab.MoveNext
Loop

For K1 = 0 To 20
    K3 = K1 * 2
    For K2 = 0 To 20
            If TDB_Nb(K1, K2) <> 0 Then
                wsExcel.Cells(K3 + 14, K2 + 2) = TDB_Moy(K1, K2)
                wsExcel.Cells(K3 + 26, K2 + 2) = TDB_Nb(K1, K2)
                TDB_Moy(K1, K2) = TDB_Moy(K1, K2) / TDB_Nb(K1, K2)
                wsExcel.Cells(K3 + 2, K2 + 2) = TDB_Moy(K1, K2)
            End If
            
           If TCR_Nb(K1, K2) <> 0 Then
                wsExcel.Cells(K3 + 15, K2 + 2) = TCR_Moy(K1, K2)
                wsExcel.Cells(K3 + 27, K2 + 2) = TCR_Nb(K1, K2)
                TCR_Moy(K1, K2) = TCR_Moy(K1, K2) / TCR_Nb(K1, K2)
                wsExcel.Cells(K3 + 3, K2 + 2) = TCR_Moy(K1, K2)
            End If
            
    Next K2
Next K1

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub

Public Sub YTP7OPH0_Export_Feuille_DATE_OPE()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSql As String
Dim wAmjMin As String, wAmjMax As String
Dim X As String, K As Long, kMax As Long, K1 As Long, K2 As Long, K3 As Long
Dim wAAAA As Integer, wMM As Integer, wJJ As Integer
Dim wRow_T As Integer
Dim S_TP7OPHCRD As Currency, S_TP7OPHCRN As Long
Dim S_TP7OPHDBD As Currency, S_TP7OPHDBN As Long
Dim T_TP7OPHCRD As Currency, T_TP7OPHCRN As Long
Dim T_TP7OPHDBD As Currency, T_TP7OPHDBN As Long
Dim arrOPE_TP7OPHCRD() As Currency, arrOPE_TP7OPHCRN() As Long
Dim arrOPE_TP7OPHDBD() As Currency, arrOPE_TP7OPHDBN() As Long
'______________________________________________

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

Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents

wsExcel.Cells(1, 1) = "Période >": wsExcel.Columns(1).ColumnWidth = 10
wsExcel.Columns(2).ColumnWidth = 15: wsExcel.Columns(2).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
wsExcel.Columns(3).ColumnWidth = 15: wsExcel.Columns(3).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"

ReDim arrOPE_TP7OPHCRD(arrOPE_Nb + 1), arrOPE_TP7OPHCRN(arrOPE_Nb + 1)
ReDim arrOPE_TP7OPHDBD(arrOPE_Nb + 1), arrOPE_TP7OPHDBN(arrOPE_Nb + 1)
wRow_T = arrOPE_Nb + 2

For K = 1 To arrOPE_Nb + 1
    wsExcel.Cells(K + 1, 1) = arrOPE(K)
    arrOPE_TP7OPHCRD(K) = 0: arrOPE_TP7OPHCRN(K) = 0
    arrOPE_TP7OPHDBD(K) = 0: arrOPE_TP7OPHDBN(K) = 0
Next K

rsYTP7OPH0_Init oldYTP7OPH0

wCol = 2: wRow = 0
S_TP7OPHCRD = 0: S_TP7OPHCRN = 0
S_TP7OPHDBD = 0: S_TP7OPHDBN = 0
T_TP7OPHCRD = 0: T_TP7OPHCRN = 0
T_TP7OPHDBD = 0: T_TP7OPHDBN = 0

'_______________________________________________________________________________________

Do While Not rsSab.EOF
    V = rsYTP7OPH0_GetBuffer(rsSab, xYTP7OPH0)
    If oldYTP7OPH0.TP7OPHDTR <> xYTP7OPH0.TP7OPHDTR Then
        If S_TP7OPHDBD <> 0 Then wsExcel.Cells(wRow, wCol) = S_TP7OPHDBD: wsExcel.Cells(wRow, wCol).Font.Color = vbRed
        If S_TP7OPHCRD <> 0 Then wsExcel.Cells(wRow, wCol + 1) = S_TP7OPHCRD
        
        T_TP7OPHCRD = T_TP7OPHCRD + S_TP7OPHCRD: T_TP7OPHCRN = T_TP7OPHCRN + S_TP7OPHCRN
        T_TP7OPHDBD = T_TP7OPHDBD + S_TP7OPHDBD: T_TP7OPHDBN = T_TP7OPHDBN + S_TP7OPHDBN
        If T_TP7OPHDBD <> 0 Then wsExcel.Cells(wRow_T, wCol) = T_TP7OPHDBD: wsExcel.Cells(wRow_T, wCol).Font.Color = vbRed
        If T_TP7OPHCRD <> 0 Then wsExcel.Cells(wRow_T, wCol + 1) = T_TP7OPHCRD
        
        arrOPE_TP7OPHCRD(arrOPE_K) = arrOPE_TP7OPHCRD(arrOPE_K) + S_TP7OPHCRD
        arrOPE_TP7OPHCRN(arrOPE_K) = arrOPE_TP7OPHCRN(arrOPE_K) + S_TP7OPHCRN
        arrOPE_TP7OPHDBD(arrOPE_K) = arrOPE_TP7OPHDBD(arrOPE_K) + S_TP7OPHDBD
        arrOPE_TP7OPHDBN(arrOPE_K) = arrOPE_TP7OPHDBN(arrOPE_K) + S_TP7OPHDBN
       
        oldYTP7OPH0 = xYTP7OPH0
        S_TP7OPHCRD = 0: S_TP7OPHCRN = 0
        S_TP7OPHDBD = 0: S_TP7OPHDBN = 0
        T_TP7OPHCRD = 0: T_TP7OPHCRN = 0
        T_TP7OPHDBD = 0: T_TP7OPHDBN = 0
        wRow = 0
        
        wCol = wCol + 2
        wsExcel.Cells(1, wCol) = "DB " & dateImp10_S(oldYTP7OPH0.TP7OPHDTR)
        wsExcel.Cells(1, wCol).Interior.Color = RGB(230, 255, 230)
        wsExcel.Columns(wCol).ColumnWidth = 15: wsExcel.Columns(wCol).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
        'wsExcel.Cells(1, wCol).Font.Color = vbRed
        wsExcel.Cells(1, wCol + 1) = "CR " & dateImp10_S(oldYTP7OPH0.TP7OPHDTR)
        wsExcel.Cells(1, wCol + 1).Interior.Color = RGB(230, 255, 230)
        wsExcel.Columns(wCol + 1).ColumnWidth = 15: wsExcel.Columns(wCol + 1).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
    Else
        If oldYTP7OPH0.TP7OPHOPE <> xYTP7OPH0.TP7OPHOPE Then
            If S_TP7OPHDBD <> 0 Then wsExcel.Cells(wRow, wCol) = S_TP7OPHDBD: wsExcel.Cells(wRow, wCol).Font.Color = vbRed
            If S_TP7OPHCRD <> 0 Then wsExcel.Cells(wRow, wCol + 1) = S_TP7OPHCRD
             
            T_TP7OPHCRD = T_TP7OPHCRD + S_TP7OPHCRD: T_TP7OPHCRN = T_TP7OPHCRN + S_TP7OPHCRN
            T_TP7OPHDBD = T_TP7OPHDBD + S_TP7OPHDBD: T_TP7OPHDBN = T_TP7OPHDBN + S_TP7OPHDBN
            
            arrOPE_TP7OPHCRD(arrOPE_K) = arrOPE_TP7OPHCRD(arrOPE_K) + S_TP7OPHCRD
            arrOPE_TP7OPHCRN(arrOPE_K) = arrOPE_TP7OPHCRN(arrOPE_K) + S_TP7OPHCRN
            arrOPE_TP7OPHDBD(arrOPE_K) = arrOPE_TP7OPHDBD(arrOPE_K) + S_TP7OPHDBD
            arrOPE_TP7OPHDBN(arrOPE_K) = arrOPE_TP7OPHDBN(arrOPE_K) + S_TP7OPHDBN
           
            oldYTP7OPH0 = xYTP7OPH0
            S_TP7OPHCRD = 0: S_TP7OPHCRN = 0
            S_TP7OPHDBD = 0: S_TP7OPHDBN = 0
            wRow = 0
            
        End If
    End If
    
    If wRow = 0 Then
        For arrOPE_K = 1 To arrOPE_Nb
            X = Trim(oldYTP7OPH0.TP7OPHOPE)
            If X = arrOPE(arrOPE_K) Then
                wRow = arrOPE_K + 1
                Exit For
            End If
        Next arrOPE_K
    End If
    S_TP7OPHDBD = S_TP7OPHDBD + oldYTP7OPH0.TP7OPHDBD
    S_TP7OPHDBN = S_TP7OPHDBN + oldYTP7OPH0.TP7OPHDBN
    S_TP7OPHCRD = S_TP7OPHCRD + oldYTP7OPH0.TP7OPHCRD
    S_TP7OPHCRN = S_TP7OPHCRN + oldYTP7OPH0.TP7OPHCRN
    rsSab.MoveNext
Loop
'_________________________________________________________________________________

If S_TP7OPHDBD <> 0 Then wsExcel.Cells(wRow, wCol) = S_TP7OPHDBD: wsExcel.Cells(wRow, wCol).Font.Color = vbRed
If S_TP7OPHCRD <> 0 Then wsExcel.Cells(wRow, wCol + 1) = S_TP7OPHCRD
T_TP7OPHCRD = T_TP7OPHCRD + S_TP7OPHCRD: T_TP7OPHCRN = T_TP7OPHCRN + S_TP7OPHCRN
T_TP7OPHDBD = T_TP7OPHDBD + S_TP7OPHDBD: T_TP7OPHDBN = T_TP7OPHDBN + S_TP7OPHDBN
If T_TP7OPHDBD <> 0 Then wsExcel.Cells(wRow_T, wCol) = T_TP7OPHDBD: wsExcel.Cells(wRow_T, wCol).Font.Color = vbRed
If T_TP7OPHCRD <> 0 Then wsExcel.Cells(wRow_T, wCol + 1) = T_TP7OPHCRD
        
arrOPE_TP7OPHCRD(arrOPE_K) = arrOPE_TP7OPHCRD(arrOPE_K) + S_TP7OPHCRD
arrOPE_TP7OPHCRN(arrOPE_K) = arrOPE_TP7OPHCRN(arrOPE_K) + S_TP7OPHCRN
arrOPE_TP7OPHDBD(arrOPE_K) = arrOPE_TP7OPHDBD(arrOPE_K) + S_TP7OPHDBD
arrOPE_TP7OPHDBN(arrOPE_K) = arrOPE_TP7OPHDBN(arrOPE_K) + S_TP7OPHDBN
'_________________________________________________________________________________
wsExcel.Cells(wRow_T, 1) = "Total"
wsExcel.Cells(1, 2) = "DB Total"
wsExcel.Cells(1, 3) = "CR Total"
For K = 1 To wRow_T
    wsExcel.Cells(K, 1).Font.Bold = True
    wsExcel.Cells(K, 1).Interior.Color = RGB(230, 255, 230)
    wsExcel.Cells(K, 2).Font.Bold = True
    wsExcel.Cells(K, 2).Interior.Color = RGB(255, 250, 250)
    wsExcel.Cells(K, 3).Font.Bold = True
    wsExcel.Cells(K, 3).Interior.Color = RGB(250, 250, 255)
    
Next K

For K = 1 To wCol + 1
    wsExcel.Cells(1, K).Font.Bold = True
    wsExcel.Cells(1, K).Interior.Color = RGB(230, 255, 230)
    wsExcel.Cells(wRow_T, K).Font.Bold = True
    wsExcel.Cells(wRow_T, K).Interior.Color = RGB(230, 255, 230)
    
Next K

For K = 1 To arrOPE_Nb + 1
    wRow = K + 1
    If arrOPE_TP7OPHDBD(K) <> 0 Then wsExcel.Cells(wRow, 2) = arrOPE_TP7OPHDBD(K): wsExcel.Cells(wRow, 2).Font.Color = vbRed
    If arrOPE_TP7OPHCRD(K) <> 0 Then wsExcel.Cells(wRow, 3) = arrOPE_TP7OPHCRD(K)
    arrOPE_TP7OPHDBD(arrOPE_Nb + 1) = arrOPE_TP7OPHDBD(arrOPE_Nb + 1) + arrOPE_TP7OPHDBD(K)
    arrOPE_TP7OPHCRD(arrOPE_Nb + 1) = arrOPE_TP7OPHCRD(arrOPE_Nb + 1) + arrOPE_TP7OPHCRD(K)
Next K

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

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

Public Sub fgdetail_Sort()
If fgDetail.Rows > 1 Then
    fgDetail.Row = 1
    fgDetail.RowSel = fgDetail.Rows - 1
    
    If fgDetail_Sort1_Old = fgDetail_Sort1 Then
        If fgDetail_SortAD = 5 Then
            fgDetail_SortAD = 6
        Else
            fgDetail_SortAD = 5
        End If
    Else
        fgDetail_SortAD = 5
    End If
    fgDetail_Sort1_Old = fgDetail_Sort1
    
    fgDetail.Col = fgDetail_Sort1
    fgDetail.ColSel = fgDetail_Sort2
    fgDetail.Sort = fgDetail_SortAD
End If

End Sub


Public Sub fgBIAMVT_Sort()
If fgBIAMVT.Rows > 1 Then
    fgBIAMVT.Row = 1
    fgBIAMVT.RowSel = fgBIAMVT.Rows - 1
    
    If fgBIAMVT_Sort1_Old = fgBIAMVT_Sort1 Then
        If fgBIAMVT_SortAD = 5 Then
            fgBIAMVT_SortAD = 6
        Else
            fgBIAMVT_SortAD = 5
        End If
    Else
        fgBIAMVT_SortAD = 5
    End If
    fgBIAMVT_Sort1_Old = fgBIAMVT_Sort1
    
    fgBIAMVT.Col = fgBIAMVT_Sort1
    fgBIAMVT.ColSel = fgBIAMVT_Sort2
    fgBIAMVT.Sort = fgBIAMVT_SortAD
End If

End Sub


Public Sub fgCPTPIE_Sort()
If fgCPTPIE.Rows > 1 Then
    fgCPTPIE.Row = 1
    fgCPTPIE.RowSel = fgCPTPIE.Rows - 1
    
    If fgCPTPIE_Sort1_Old = fgCPTPIE_Sort1 Then
        If fgCPTPIE_SortAD = 5 Then
            fgCPTPIE_SortAD = 6
        Else
            fgCPTPIE_SortAD = 5
        End If
    Else
        fgCPTPIE_SortAD = 5
    End If
    fgCPTPIE_Sort1_Old = fgCPTPIE_Sort1
    
    fgCPTPIE.Col = fgCPTPIE_Sort1
    fgCPTPIE.ColSel = fgCPTPIE_Sort2
    fgCPTPIE.Sort = fgCPTPIE_SortAD
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

wFct = UCase$(Trim(mId$(Msg, 1, 12)))
Call BiaPgmAut_Init(wFct, YTP7OPH0_Aut)

'blnSetfocus = True
Form_Init


Select Case wFct
    Case "@SAB_ICC": blnAuto = True
    Case Else: blnAuto = False
End Select

End Sub


Public Sub Form_Init()
Dim V, xSql As String, X As String
Dim K As Long

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True


cmdReset
blnControl = False

fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False


fraSelect_Options_1.BorderStyle = 0
lblTP7OPHCOM.ForeColor = vbMagenta

lstW.Visible = False

fgDetail.Visible = False: fraDetail.Visible = False
fgDetail_FormatString = fgDetail.FormatString
Call DTPicker_Set(txtSelect_TP7OPHDTR_Max, YBIATAB0_DATE_CPT_J) '
Call DTPicker_Set(txtSelect_TP7OPHDTR_Min, YBIATAB0_DATE_CPT_JP0) '

fgBIAMVT.Visible = False
fgBIAMVT_FormatString = fgBIAMVT.FormatString

fgCPTPIE.Visible = False
fgCPTPIE_FormatString = fgCPTPIE.FormatString

cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1  - sélection (filtre)"
cboSelect_SQL.AddItem "En  - export Nostro (2006-2010)"
cboSelect_SQL.AddItem "Ek  - export Caisse (2006-2010)"
cboSelect_SQL.AddItem "Ep  - export Période"
cboSelect_SQL.ListIndex = 0


lstW.Clear


'Initialisation opération________________________________________________________________________________
arrOPE_Nb = 0
ReDim Preserve arrOPE(1000)

cboSelect_TP7OPHOPE.Clear
cboSelect_TP7OPHOPE.AddItem ""
xSql = "select distinct TP7OPHOPE from " & paramIBM_Library_SABSPE & ".YTP7OPH0 order by TP7OPHOPE"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    arrOPE_Nb = arrOPE_Nb + 1
    arrOPE(arrOPE_Nb) = Trim(rsSab("TP7OPHOPE"))
    cboSelect_TP7OPHOPE.AddItem Trim(rsSab("TP7OPHOPE"))
    rsSab.MoveNext
Loop
ReDim Preserve arrOPE(arrOPE_Nb + 1)

'Initialisation devise________________________________________________________________________________
arrDev_Nb = 0
ReDim Preserve arrDev(1000)

cboSelect_TP7OPHDEV.Clear
cboSelect_TP7OPHDEV.AddItem ""
xSql = "select distinct TP7OPHDEV from " & paramIBM_Library_SABSPE & ".YTP7OPH0 order by TP7OPHDEV"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    arrDev_Nb = arrDev_Nb + 1
    arrDev(arrDev_Nb) = Trim(rsSab("TP7OPHDEV"))
    cboSelect_TP7OPHDEV.AddItem Trim(rsSab("TP7OPHDEV"))
    rsSab.MoveNext
Loop
ReDim Preserve arrDev(arrDev_Nb + 1)

'Initialisation Compte________________________________________________________________________________

cboSelect_TP7OPHCOM.Clear
cboSelect_TP7OPHCOM.AddItem ""
xSql = "select distinct TP7OPHCOM from " & paramIBM_Library_SABSPE & ".YTP7OPH0 order by TP7OPHCOM"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_TP7OPHCOM.AddItem Trim(rsSab("TP7OPHCOM"))
    rsSab.MoveNext
Loop

fraSelect_Options.Visible = True
blnControl = True

Me.Enabled = True
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
fgDetail.Visible = False: fraDetail.Visible = False
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
fgDetail.Visible = True: fraDetail.Visible = True
End Sub

Public Sub fgBIAMVT_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgBIAMVT.Visible = False: fraDetail.Visible = False
mRow = fgBIAMVT.Row

If lRow > 0 And lRow < fgBIAMVT.Rows Then
    fgBIAMVT.Row = lRow
    For I = fgBIAMVT_arrIndex To fgBIAMVT.FixedCols Step -1
        fgBIAMVT.Col = I: fgBIAMVT.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgBIAMVT.Row = mRow
    If fgBIAMVT.Row > 0 Then
        lRow = fgBIAMVT.Row
        lColor_Old = fgBIAMVT.CellBackColor
        For I = fgBIAMVT_arrIndex To fgBIAMVT.FixedCols Step -1
          fgBIAMVT.Col = I: fgBIAMVT.CellBackColor = lColor
        Next I
    End If
End If
fgBIAMVT.LeftCol = fgBIAMVT.FixedCols
fgBIAMVT.Visible = True: fraDetail.Visible = True
End Sub

Public Sub fgCPTPIE_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgCPTPIE.Visible = False: fraDetail.Visible = False
mRow = fgCPTPIE.Row

If lRow > 0 And lRow < fgCPTPIE.Rows Then
    fgCPTPIE.Row = lRow
    For I = fgCPTPIE_arrIndex To fgCPTPIE.FixedCols Step -1
        fgCPTPIE.Col = I: fgCPTPIE.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgCPTPIE.Row = mRow
    If fgCPTPIE.Row > 0 Then
        lRow = fgCPTPIE.Row
        lColor_Old = fgCPTPIE.CellBackColor
        For I = fgCPTPIE_arrIndex To fgCPTPIE.FixedCols Step -1
          fgCPTPIE.Col = I: fgCPTPIE.CellBackColor = lColor
        Next I
    End If
End If
fgCPTPIE.LeftCol = fgCPTPIE.FixedCols
fgCPTPIE.Visible = True: fraDetail.Visible = True
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






Private Sub cboSelect_SQL_Click()
cmdSelect_Reset

End Sub


Private Sub cboSelect_TP7OPHCOM_Change()
cmdSelect_Reset

End Sub

Private Sub cboSelect_TP7OPHCOM_Click()
cmdSelect_Reset

End Sub

Private Sub cboSelect_TP7OPHDEV_Change()
cmdSelect_Reset

End Sub

Private Sub cboSelect_TP7OPHDEV_Click()
cmdSelect_Reset

End Sub

Private Sub cboSelect_TP7OPHOPE_Click()
cmdDetail_Reset
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

Private Sub cmdPrint_YTP7OPH0(blnDetail As Boolean)
Dim X As String, xSql As String, I As Integer, K As Integer
Dim wAMJ As String, xWhere As String
Dim soldeD As typeYTP7OPH0, soldeF As typeYTP7OPH0, total As typeYTP7OPH0
Dim blnXprt_Line As Boolean



Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> TP7_OPH_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Reset
fgSelect.Visible = False
'fraSelect_Options.Visible = False

Select Case cmdSelect_SQL_K
    Case "1": fraSelect_Options.Visible = True: cmdSelect_SQL_1
    Case "En": YTP7OPH0_Export
    Case "Ek": YTP7OPH0_Export
    Case "Ep": YTP7OPH0_Export
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< TP7_OPH_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus

End Sub


Private Sub fgBIAMVT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

If Y <= fgBIAMVT.RowHeightMin Then
Else
    If fgBIAMVT.Rows > 1 Then
        Call fgBIAMVT_Color(fgBIAMVT_RowClick, MouseMoveUsr.BackColor, fgBIAMVT_ColorClick)
        fgBIAMVT.Col = fgBIAMVT_arrIndex:  arrYBIAMVTH_Index = CLng(fgBIAMVT.Text)
        oldYBIAMVTH = arrYBIAMVTH(arrYBIAMVTH_Index)
        xYBIAMVTH = oldYBIAMVTH
        fgCPTPIE_Display

   End If
End If


End Sub


Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next


If Y <= fgDetail.RowHeightMin Then
Else
    If fgDetail.Rows > 1 Then
        Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
        fgDetail.Col = fgDetail_arrIndex:  arrYTP7OPH0_Index = CLng(fgDetail.Text)
        oldYTP7OPH0 = arrYTP7OPH0(arrYTP7OPH0_Index)
        xYTP7OPH0 = oldYTP7OPH0
        fgBIAMVT_Display

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
    Case Is = 27: cmdContext_Quit
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select


End Sub

Public Sub cmdContext_Quit()
'blnControl = False
lstErr.Clear: lstErr.Height = 200

If fgCPTPIE.Visible Then
    fgCPTPIE.Visible = False
    Exit Sub
End If

If fgBIAMVT.Visible Then
    fgBIAMVT.Visible = False
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
    Exit Sub

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


If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  arrYBIACPT0_Index = CLng(fgSelect.Text)
        
    oldYBIACPT0 = arrYBIACPT0(arrYBIACPT0_Index)
    xYBIACPT0 = oldYBIACPT0
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




Public Sub fgBIAMVT_Reset()
fgBIAMVT.Clear
fgBIAMVT_Sort1 = 0: fgBIAMVT_Sort2 = 0
fgBIAMVT_Sort1_Old = -1
fgBIAMVT_RowDisplay = 0: fgBIAMVT_RowClick = 0
fgBIAMVT_arrIndex = fgBIAMVT.Cols - 1
blnfgBIAMVT_DisplayLine = False
fgBIAMVT_SortAD = 6
fgBIAMVT.LeftCol = fgBIAMVT.FixedCols

End Sub

Public Sub fgCPTPIE_Reset()
fgCPTPIE.Clear
fgCPTPIE_Sort1 = 0: fgCPTPIE_Sort2 = 0
fgCPTPIE_Sort1_Old = -1
fgCPTPIE_RowDisplay = 0: fgCPTPIE_RowClick = 0
fgCPTPIE_arrIndex = fgCPTPIE.Cols - 1
blnfgCPTPIE_DisplayLine = False
fgCPTPIE_SortAD = 6
fgCPTPIE.LeftCol = fgCPTPIE.FixedCols

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







Private Sub mnuPrint_Detail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdPrint_YTP7OPH0 True

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint_Recap_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdPrint_YTP7OPH0 False

Me.Enabled = True: Me.MousePointer = 0
End Sub




















Private Sub txtSelect_TP7OPHDTR_Max_Change()
cmdDetail_Reset

End Sub

Private Sub txtSelect_TP7OPHDTR_Max_Click()
cmdDetail_Reset

End Sub

Private Sub txtSelect_TP7OPHDTR_Min_Change()
cmdDetail_Reset

End Sub

Private Sub txtSelect_TP7OPHDTR_Min_Click()
cmdDetail_Reset

End Sub

