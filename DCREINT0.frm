VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDCREINT0 
   AutoRedraw      =   -1  'True
   Caption         =   "DCREINT0 : maintenance"
   ClientHeight    =   9315
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
   Icon            =   "DCREINT0.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9315
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
      Height          =   8688
      Left            =   -48
      TabIndex        =   2
      Top             =   600
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   15319
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "DCREINT0"
      TabPicture(0)   =   "DCREINT0.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "......."
      TabPicture(1)   =   "DCREINT0.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDetail"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraDetail 
         BackColor       =   &H00A0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6852
         Left            =   -71640
         TabIndex        =   17
         Top             =   960
         Width           =   7500
         Begin VB.CommandButton cmdDetail_Update 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Enregistrer"
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
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   6150
            Width           =   1095
         End
         Begin VB.CommandButton cmdDetail_Delete 
            BackColor       =   &H000000FF&
            Caption         =   "Supprimer"
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
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   6150
            Width           =   1095
         End
         Begin VB.CommandButton cmdDetail_Copy 
            BackColor       =   &H00FF80FF&
            Caption         =   "Copier"
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
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   6150
            Width           =   1095
         End
         Begin VB.CommandButton cmdDetail_Quit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Abandonner"
            Enabled         =   0   'False
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
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   6150
            Width           =   1095
         End
         Begin VB.CommandButton cmdDetail_Control 
            BackColor       =   &H0080C0FF&
            Caption         =   "Saisie"
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
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   6120
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   5865
            Left            =   135
            TabIndex        =   23
            Top             =   100
            Width           =   7200
            _ExtentX        =   12700
            _ExtentY        =   10345
            _Version        =   393216
            Rows            =   1
            Cols            =   5
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   16316664
            ForeColor       =   8388608
            BackColorFixed  =   10543359
            ForeColorFixed  =   0
            BackColorSel    =   12648384
            BackColorBkg    =   16777215
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            ScrollBars      =   2
            AllowUserResizing=   3
            FormatString    =   "*    |<Champ          |< Libellé                                    |> Valeur                     |>Modification            "
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
         Height          =   8232
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   13296
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
            Height          =   5640
            Left            =   6600
            TabIndex        =   27
            Top             =   1800
            Visible         =   0   'False
            Width           =   6492
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
            Left            =   10200
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   240
            Width           =   2772
         End
         Begin VB.TextBox txtDetail 
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   9360
            TabIndex        =   12
            Top             =   840
            Visible         =   0   'False
            Width           =   1452
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Exécuter le traitement"
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
            Left            =   10920
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   720
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
            Height          =   1005
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   8832
            Begin VB.TextBox txtSelect_CREINTDOS 
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   324
               Left            =   7200
               TabIndex        =   26
               Top             =   480
               Width           =   1212
            End
            Begin VB.ComboBox txtSelect_CREINTDEV 
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
               Left            =   4080
               Sorted          =   -1  'True
               TabIndex        =   14
               Text            =   "txtSelect_GAPPISDEV"
               Top             =   500
               Width           =   852
            End
            Begin VB.ComboBox txtSelect_CREINTNAT 
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
               Left            =   5640
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   480
               Width           =   1332
            End
            Begin VB.TextBox txtSelect_CREINTCLI 
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   324
               Left            =   2280
               TabIndex        =   11
               Top             =   500
               Width           =   1212
            End
            Begin MSComCtl2.DTPicker txtSelect_CREINTPER 
               Height          =   300
               Left            =   480
               TabIndex        =   7
               Top             =   500
               Width           =   1332
               _ExtentX        =   2355
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
               Format          =   54853635
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_CREINTDOS 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Dossier"
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
               Left            =   7440
               TabIndex        =   25
               Top             =   240
               Width           =   852
            End
            Begin VB.Label lblSelect_CREINTDEV 
               BackColor       =   &H00F0FFFF&
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
               Left            =   4200
               TabIndex        =   16
               Top             =   204
               Width           =   612
            End
            Begin VB.Label lblSelect_CREINTNAT 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Nature"
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
               Left            =   5760
               TabIndex        =   15
               Top             =   204
               Width           =   852
            End
            Begin VB.Label lblSelect_CREINTCLI 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Client"
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
               Left            =   2520
               TabIndex        =   10
               Top             =   204
               Width           =   612
            End
            Begin VB.Label lblSelect_CREINTPER 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Période"
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
               TabIndex        =   9
               Top             =   200
               Width           =   612
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6828
            Left            =   240
            TabIndex        =   5
            Top             =   1320
            Width           =   12912
            _ExtentX        =   22781
            _ExtentY        =   12039
            _Version        =   393216
            Rows            =   1
            Cols            =   13
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   -2147483633
            ForeColor       =   12582912
            BackColorFixed  =   15794175
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   -2147483633
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"DCREINT0.frx":0342
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
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
      Picture         =   "DCREINT0.frx":0415
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
   Begin VB.Menu mnufgSelect 
      Caption         =   "mnufgSelect"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmDCREINT0"
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
Dim DCREINT0_Aut As typeAuthorization
Dim blnAuto As Boolean, blnError As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean


Dim fgDetail_FormatString As String, fgDetail_K As Integer
Dim fgDetail_RowDisplay As Integer, fgDetail_RowClick As Integer, fgDetail_ColClick As Integer
Dim fgDetail_ColorClick As Long, fgDetail_ColorDisplay As Long
Dim fgDetail_Sort1 As Integer, fgDetail_Sort2 As Integer
Dim fgDetail_SortAD As Integer, fgDetail_Sort1_Old As Integer
Dim fgDetail_arrIndex As Integer
Dim blnfgDetail_DisplayLine As Boolean

'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
Dim xDCREINT0 As typeDCREINT0, newDCREINT0 As typeDCREINT0, oldDCREINT0 As typeDCREINT0
Dim arrDCREINT0() As typeDCREINT0, arrDCREINT0_Nb As Long, arrDCREINT0_Max As Long, arrDCREINT0_Index As Long

Dim txtDetail_Type As String, txtDetail_Update As String
Dim txtDetail_Field As String
Dim txtDetail_blnUpdate As Boolean, txtDetail_ColorUpdate As Long
Dim txtDetail_Row As Integer

Dim cmdSelect_SQL_K As String
Dim arrTRIM(21) As typeDCRETA
Dim arrDCRETA(500) As typeDCRETA, arrDCRETA_Nb As Integer


Dim wFolder_CREGS601P1 As String
Public Sub fgDetail_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgDetail.Row

If lRow > 0 And lRow < fgDetail.Rows Then
    fgDetail.Row = lRow
    For I = 0 To fgDetail_arrIndex
        fgDetail.Col = I: fgDetail.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgDetail.Row = mRow
    If fgDetail.Row > 0 Then
        lRow = fgDetail.Row
        lColor_Old = fgDetail.CellBackColor
        For I = 0 To fgDetail_arrIndex
          fgDetail.Col = I: fgDetail.CellBackColor = lColor
        Next I
        fgDetail.Col = 0
    End If
End If

End Sub
Private Sub fgDetail_Display()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = "fgDetail_Display"
SSTab1.Tab = 0
fraDetail.Visible = False
cmdDetail_Delete.Visible = DCREINT0_Aut.Saisir
cmdDetail_Copy.Visible = DCREINT0_Aut.Saisir
cmdDetail_Update.Visible = False
cmdDetail_Quit.Enabled = True
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString

    
For I = 1 To 62
         
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_DisplayLine I
Next I

fraDetail.Visible = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Public Sub fgDetail_DisplayLine(lIndex As Long)
Dim K As Integer, wField As String
On Error Resume Next



wField = Trim(Mid$(ddsDCREINT0(lIndex), 4, 10))
fgDetail.Col = 0: fgDetail.Text = Mid$(ddsDCREINT0(lIndex), 1, 2)
fgDetail.Col = 1: fgDetail.Text = wField
fgDetail.Col = 2: fgDetail.Text = Mid$(ddsDCREINT0(lIndex), 15, 30)
fgDetail.Col = 3
Select Case wField
    Case "CREINTSTA": fgDetail.Text = oldDCREINT0.CREINTSTA
    Case "CREINTVER": fgDetail.Text = oldDCREINT0.CREINTVER
    Case "CREINTPER": fgDetail.Text = oldDCREINT0.CREINTPER
    Case "CREINTETA": fgDetail.Text = oldDCREINT0.CREINTETA
    Case "CREINTAGE": fgDetail.Text = oldDCREINT0.CREINTAGE
    Case "CREINTSER": fgDetail.Text = oldDCREINT0.CREINTSER
    Case "CREINTSSE": fgDetail.Text = oldDCREINT0.CREINTSSE
    Case "CREINTDOS": fgDetail.Text = oldDCREINT0.CREINTDOS
    Case "CREINTPRE": fgDetail.Text = oldDCREINT0.CREINTPRE
    Case "CREINTNAT": fgDetail.Text = oldDCREINT0.CREINTNAT
    
    Case "CREINTNAP": fgDetail.Text = oldDCREINT0.CREINTNAP
    Case "CREINTCLI": fgDetail.Text = oldDCREINT0.CREINTCLI
    Case "CREINTMT0": fgDetail.Text = Format$(oldDCREINT0.CREINTMT0, "### ### ### ###.00")
    Case "CREINTDEV": fgDetail.Text = oldDCREINT0.CREINTDEV
    Case "CREINTECH": fgDetail.Text = oldDCREINT0.CREINTECH
    Case "CREINTMTX": fgDetail.Text = Format$(oldDCREINT0.CREINTMTX, "### ### ### ###.00")
    Case "CREINTTOF": fgDetail.Text = Format$(oldDCREINT0.CREINTTOF, "###.00000")
    Case "CREINTTOM": fgDetail.Text = Format$(oldDCREINT0.CREINTTOM, "###.00000")
    Case "CREINTPERK": fgDetail.Text = oldDCREINT0.CREINTPERK
    Case "CREINTPERN": fgDetail.Text = oldDCREINT0.CREINTPERN
    Case "CREINTUAMJ": fgDetail.Text = oldDCREINT0.CREINTUAMJ
    Case "CREINTUHMS": fgDetail.Text = oldDCREINT0.CREINTUHMS
    
    Case "CREINTT01": fgDetail.Text = Format$(oldDCREINT0.CREINTT01, "### ### ### ###.00")
    Case "CREINTM01": fgDetail.Text = Format$(oldDCREINT0.CREINTM01, "### ### ### ###.00")
    Case "CREINTT02": fgDetail.Text = Format$(oldDCREINT0.CREINTT02, "### ### ### ###.00")
    Case "CREINTM02": fgDetail.Text = Format$(oldDCREINT0.CREINTM02, "### ### ### ###.00")
    Case "CREINTT03": fgDetail.Text = Format$(oldDCREINT0.CREINTT03, "### ### ### ###.00")
    Case "CREINTM03": fgDetail.Text = Format$(oldDCREINT0.CREINTM03, "### ### ### ###.00")
    Case "CREINTT04": fgDetail.Text = Format$(oldDCREINT0.CREINTT04, "### ### ### ###.00")
    Case "CREINTM04": fgDetail.Text = Format$(oldDCREINT0.CREINTM04, "### ### ### ###.00")
    
    Case "CREINTT11": fgDetail.Text = Format$(oldDCREINT0.CREINTT11, "### ### ### ###.00")
    Case "CREINTM11": fgDetail.Text = Format$(oldDCREINT0.CREINTM11, "### ### ### ###.00")
    Case "CREINTT12": fgDetail.Text = Format$(oldDCREINT0.CREINTT12, "### ### ### ###.00")
    Case "CREINTM12": fgDetail.Text = Format$(oldDCREINT0.CREINTM12, "### ### ### ###.00")
    Case "CREINTT13": fgDetail.Text = Format$(oldDCREINT0.CREINTT13, "### ### ### ###.00")
    Case "CREINTM13": fgDetail.Text = Format$(oldDCREINT0.CREINTM13, "### ### ### ###.00")
    Case "CREINTT14": fgDetail.Text = Format$(oldDCREINT0.CREINTT14, "### ### ### ###.00")
    Case "CREINTM14": fgDetail.Text = Format$(oldDCREINT0.CREINTM14, "### ### ### ###.00")
    
    Case "CREINTT21": fgDetail.Text = Format$(oldDCREINT0.CREINTT21, "### ### ### ###.00")
    Case "CREINTM21": fgDetail.Text = Format$(oldDCREINT0.CREINTM21, "### ### ### ###.00")
    Case "CREINTT22": fgDetail.Text = Format$(oldDCREINT0.CREINTT22, "### ### ### ###.00")
    Case "CREINTM22": fgDetail.Text = Format$(oldDCREINT0.CREINTM22, "### ### ### ###.00")
    Case "CREINTT23": fgDetail.Text = Format$(oldDCREINT0.CREINTT23, "### ### ### ###.00")
    Case "CREINTM23": fgDetail.Text = Format$(oldDCREINT0.CREINTM23, "### ### ### ###.00")
    Case "CREINTT24": fgDetail.Text = Format$(oldDCREINT0.CREINTT24, "### ### ### ###.00")
    Case "CREINTM24": fgDetail.Text = Format$(oldDCREINT0.CREINTM24, "### ### ### ###.00")
    
    Case "CREINTT31": fgDetail.Text = Format$(oldDCREINT0.CREINTT31, "### ### ### ###.00")
    Case "CREINTM31": fgDetail.Text = Format$(oldDCREINT0.CREINTM31, "### ### ### ###.00")
    Case "CREINTT32": fgDetail.Text = Format$(oldDCREINT0.CREINTT32, "### ### ### ###.00")
    Case "CREINTM32": fgDetail.Text = Format$(oldDCREINT0.CREINTM32, "### ### ### ###.00")
    Case "CREINTT33": fgDetail.Text = Format$(oldDCREINT0.CREINTT33, "### ### ### ###.00")
    Case "CREINTM33": fgDetail.Text = Format$(oldDCREINT0.CREINTM33, "### ### ### ###.00")
    Case "CREINTT34": fgDetail.Text = Format$(oldDCREINT0.CREINTT34, "### ### ### ###.00")
    Case "CREINTM34": fgDetail.Text = Format$(oldDCREINT0.CREINTM34, "### ### ### ###.00")
    
    Case "CREINTT41": fgDetail.Text = Format$(oldDCREINT0.CREINTT41, "### ### ### ###.00")
    Case "CREINTM41": fgDetail.Text = Format$(oldDCREINT0.CREINTM41, "### ### ### ###.00")
    Case "CREINTT42": fgDetail.Text = Format$(oldDCREINT0.CREINTT42, "### ### ### ###.00")
    Case "CREINTM42": fgDetail.Text = Format$(oldDCREINT0.CREINTM42, "### ### ### ###.00")
    Case "CREINTT43": fgDetail.Text = Format$(oldDCREINT0.CREINTT43, "### ### ### ###.00")
    Case "CREINTM43": fgDetail.Text = Format$(oldDCREINT0.CREINTM43, "### ### ### ###.00")
    Case "CREINTT44": fgDetail.Text = Format$(oldDCREINT0.CREINTT44, "### ### ### ###.00")
    Case "CREINTM44": fgDetail.Text = Format$(oldDCREINT0.CREINTM44, "### ### ### ###.00")
    

End Select
If Mid$(ddsDCREINT0(lIndex), 2, 1) <> "*" Then fgDetail.CellBackColor = &HC0FFFF
'fgDetail.Col = fgDetail_arrIndex: fgDetail.Text = lIndex
End Sub

Public Sub fgDetail_Reset()
fgDetail.Clear
fgDetail_Sort1 = 0: fgDetail_Sort2 = 0
fgDetail_Sort1_Old = -1
fgDetail_RowDisplay = 0: fgDetail_RowClick = 0
fgDetail_arrIndex = fgDetail.Cols - 1
blnfgDetail_DisplayLine = False
fgDetail_SortAD = 6
fgDetail.LeftCol = 0

End Sub


Public Sub fgDetail_Sort()
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
Public Sub fgDetail_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgDetail.Rows - 1
    fgDetail.Row = I
    fgDetail.Col = fgDetail_arrIndex ':  fgdetail_Index = CLng(fgDetail.Text)
    Select Case lK
       ' Case 0: X = arrYPDCMVT0(arrYPDCMVT0_Index).PDCMVTDTR
    End Select
    fgDetail.Col = fgDetail_arrIndex - 1
    fgDetail.Text = X
Next I


fgDetail_Sort1 = fgDetail_arrIndex - 1: fgDetail_Sort2 = fgDetail_arrIndex - 1
fgDetail_Sort
End Sub



Public Sub cmdSelect_Reset()
If blnControl Then
    lstErr.Clear
    lstW.Visible = False
    fgSelect.Visible = False
    fraDetail.Visible = False
    txtDetail.Visible = False
    cmdSelect_Ok.Visible = True
    cmdSelect_SQL_K = Trim(Mid$(cboSelect_SQL, 1, 2))
    Select Case cmdSelect_SQL_K
        Case "2": cmdSelect_Ok.Visible = False: cmdSelect_SQL_2
        Case "I": cmdSelect_SQL_Import_TA_Init
    End Select
End If

End Sub


Private Sub cmdSelect_SQL()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim wNum As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdDCREINT0_SQL"
blnOk = False
   
Call DTPicker_Control(txtSelect_CREINTPER, wAmjMax)
xWhere = " where CREINTPER = " & wAmjMax

wNum = Val(txtSelect_CREINTCLI)
If wNum <> 0 Then blnOk = True: xWhere = xWhere & " and CREINTCLI = " & wNum

wNum = Val(txtSelect_CREINTDOS)
If wNum <> 0 Then blnOk = True: xWhere = xWhere & " and CREINTDOS = " & wNum


X = Trim(txtSelect_CREINTDEV)
If X <> "" Then blnOk = True: xWhere = xWhere & " and CREINTDEV = '" & X & "'"

X = Trim(txtSelect_CREINTNAT)
If X <> "" Then blnOk = True: xWhere = xWhere & " and CREINTNAT = '" & X & "'"

'If Not blnOk Then
'    Call MsgBox("préciser plus de critères : n°client, code opération...", vbCritical, "BIA_DWH : DCREINT0")
'    Exit Sub
'End If

arrDCREINT0_SQL xWhere
fgSelect_Display



Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
    


End Sub
Private Sub cmdSelect_SQL_2()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim wNum As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdDCREINT0_SQL_2"
blnOk = False
   
Call DTPicker_Control(txtSelect_CREINTPER, wAmjMax)
xWhere = " where CREINTPER = " & wAmjMax & " and CREINTUAMJ = 0"


arrDCREINT0_SQL xWhere
fgSelect_Display



Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
    


End Sub

Private Sub arrDCREINT0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrDCREINT0(101)
arrDCREINT0_Max = 100: arrDCREINT0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_BODWH & ".DCREINT0 " & xWhere & " order by CREINTDOS , CREINTPRE "
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsDCREINT0_GetBuffer(rsSab, xDCREINT0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDCREINT0.fgselect_Display"
        '' Exit Sub
     Else
         arrDCREINT0_Nb = arrDCREINT0_Nb + 1
         If arrDCREINT0_Nb > arrDCREINT0_Max Then
             arrDCREINT0_Max = arrDCREINT0_Max + 100
             ReDim Preserve arrDCREINT0(arrDCREINT0_Max)
         End If
         
         arrDCREINT0(arrDCREINT0_Nb) = xDCREINT0
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



'______________________________________________________________________

Private Sub fgSelect_Display()

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgSelect_Display"
    
For I = 1 To arrDCREINT0_Nb
         
    xDCREINT0 = arrDCREINT0(I)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine I
Next I

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrDCREINT0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
On Error Resume Next


fgSelect.Col = 0: fgSelect.Text = xDCREINT0.CREINTDOS & "-" & Format(xDCREINT0.CREINTPRE, "000")
fgSelect.Col = 1: fgSelect.Text = xDCREINT0.CREINTNAT & " " & xDCREINT0.CREINTNAP
fgSelect.Col = 2: fgSelect.Text = xDCREINT0.CREINTCLI
fgSelect.Col = 3: fgSelect.Text = Format$(xDCREINT0.CREINTMT0, "### ### ### ##0.00")
fgSelect.Col = 4: fgSelect.Text = xDCREINT0.CREINTDEV
fgSelect.Col = 5: fgSelect.Text = Format$(xDCREINT0.CREINTMTX, "### ### ### ##0.00")
fgSelect.Col = 6: fgSelect.Text = dateImp10(xDCREINT0.CREINTECH)
fgSelect.Col = 7: fgSelect.Text = Format$(xDCREINT0.CREINTTOF, "##0.00000")
fgSelect.Col = 8: fgSelect.Text = Format$(xDCREINT0.CREINTTOM, "##0.00000")
fgSelect.Col = 9: fgSelect.Text = xDCREINT0.CREINTPERK & " " & xDCREINT0.CREINTPERN
If xDCREINT0.CREINTUAMJ > 0 Then fgSelect.Col = 10: fgSelect.Text = dateImp10(xDCREINT0.CREINTUAMJ) & " " & timeImp8(xDCREINT0.CREINTUHMS)


fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
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
    fgSelect.Col = lK
    X = Format$(Val(fgSelect.Text), "0000000")
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
    'Select Case lK
    '    Case 1, 2: fgSelect.Text = X
    'End Select
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
Call BiaPgmAut_Init(wFct, DCREINT0_Aut)

'blnSetfocus = True
Form_Init

Select Case wFct
    Case Else: blnAuto = False
End Select





End Sub


Public Sub Form_Init()
Dim V, xSql As String

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

lstErr.Visible = True

blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False


fgDetail_FormatString = fgDetail.FormatString
fraDetail.Visible = False
Set fraDetail.Container = fraTab0
cmdReset

fraDetail.Left = 5800 '6300
fraDetail.Top = fgSelect.Top
Set txtDetail.Container = fraDetail

txtDetail.BackColor = &HC0FFC0
txtDetail_ColorUpdate = &H80C0FF
wAmjMax = dateFinDeMois(YBIATAB0_DATE_CPT_J)
If wAmjMax > YBIATAB0_DATE_CPT_J Then wAmjMax = dateElp("FinDeMoisP", 0, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtSelect_CREINTPER, wAmjMax) '

ddsCREINT0_Init


txtSelect_CREINTNAT.Clear
txtSelect_CREINTNAT.AddItem " "
xSql = "select distinct CREINTNAT from " & paramIBM_Library_BODWH & ".DCREINT0 order by CREINTNAT"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    txtSelect_CREINTNAT.AddItem Trim(rsSab("CREINTNAT"))
    rsSab.MoveNext
Loop

txtSelect_CREINTDEV.Clear
txtSelect_CREINTDEV.AddItem " "
xSql = "select distinct CREINTDEV from " & paramIBM_Library_BODWH & ".DCREINT0 order by CREINTDEV"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    txtSelect_CREINTDEV.AddItem Trim(rsSab("CREINTDEV"))
    rsSab.MoveNext
Loop

cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1  - sélection (filtre)"
cboSelect_SQL.AddItem "2  - Dossiers sans Tableau d'amortissement"
If DCREINT0_Aut.Saisir Then
    cboSelect_SQL.AddItem "I  - Import Tableau Amortissement"
End If
cboSelect_SQL.ListIndex = 0

'Call MsgBox("- Les contrôles sur les champs modifiés sont très succints," & vbCrLf _
'           & "- la cohérence des données n'est pas assûrée," & vbCrLf _
'           & "- il n'y a pas de trace des modifications." & vbCrLf & vbCrLf _
'           & "Les modifications du fichier DCREINT0 sont sous votre responsabilité !" _
'           , vbExclamation, "maintenance du fichier BODWH / DCREINT0")
           
           

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
mRow = fgSelect.Row
fgSelect.Visible = False
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
        For I = 0 To fgSelect_arrIndex
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
    End If
End If
fgSelect.LeftCol = 0
fgSelect.Visible = True
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


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdDetail_Control_Click()
fraDetail_Control

End Sub

Private Sub cmdDetail_Copy_Click()
Dim blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass
blnOk = False
xDCREINT0 = newDCREINT0
xDCREINT0.CREINTDOS = 900000000
Do
    xDCREINT0.CREINTDOS = xDCREINT0.CREINTDOS + 1
      
    If Not IsNull(sqlDCREINT0_Read(xDCREINT0)) Then blnOk = True
Loop Until blnOk
newDCREINT0.CREINTDOS = xDCREINT0.CREINTDOS

V = sqlDCREINT0_Insert(newDCREINT0)
fraDetail.Visible = False

cmdSelect_SQL
Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdDetail_Delete_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

V = sqlDCREINT0_Delete(oldDCREINT0)
fraDetail.Visible = False

cmdSelect_SQL
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdDetail_Quit_Click()
fraDetail.Visible = False
End Sub

Private Sub cmdDetail_Update_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
'If IsNull(cmdDCREINT0_Control) Then
    cmdDCREINT0_Update
    fraDetail.Visible = False
    cmdSelect_SQL
'End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdDCREINT0_Update()

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

V = sqlDCREINT0_Update(newDCREINT0, oldDCREINT0)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

If Not IsNull(V) Then MsgBox V, vbCritical

End Sub

Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_DWH_cmdSelect_Ok ........"): DoEvents
    
    
 Select Case cmdSelect_SQL_K
    Case "1": fraSelect_Options.Visible = True: cmdSelect_SQL
    Case "I": cmdSelect_SQL_Import_TA
End Select
   
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_DWH_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

If fgDetail.Rows > 1 Then
    'Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
    fgDetail_RowClick = fgDetail.Row
    If txtDetail_blnUpdate Then fraDetail_Control
    txtDetail.Visible = False
    'txtDetail = ""
    
    fgDetail.Col = 0
    
    txtDetail_Type = Mid$(fgDetail.Text, 1, 1)
    txtDetail_Update = Mid$(fgDetail.Text, 2, 1)
    If txtDetail_Update <> "*" And DCREINT0_Aut.Saisir Then
        txtDetail_Row = fgDetail.Row
        fgDetail.Col = 1
        txtDetail_Field = Trim(fgDetail.Text)
        fgDetail.Col = 4
        txtDetail.Top = fgDetail.Top + fgDetail.CellTop
        txtDetail.Left = fgDetail.Left + fgDetail.CellLeft
       If fgDetail.CellBackColor <> txtDetail_ColorUpdate Then fgDetail.Col = 3
        
        txtDetail = Trim(fgDetail.Text)
        txtDetail_blnUpdate = False: cmdDetail_Control.Visible = True
        cmdDetail_Update.Visible = False

        txtDetail.Visible = True
        txtDetail.SetFocus
    End If
End If
End Sub


Private Sub fgDetail_Scroll()
txtDetail_blnUpdate = False
txtDetail.Visible = False
cmdDetail_Control.Visible = False
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
blnControl = False
lstErr.Clear: lstErr.Height = 200

If txtDetail.Visible Then
    txtDetail.Visible = False
    Exit Sub
End If
If fraDetail.Visible Then
    fraDetail.Visible = False
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
    If Not fgSelect.Version Then cmdSelect_Ok_Click
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





Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wOrigine As String
On Error Resume Next

txtDetail.Visible = False

If y <= fgSelect.RowHeightMin Then
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
        fgSelect.Col = fgSelect_arrIndex:  arrDCREINT0_Index = CLng(fgSelect.Text)
        oldDCREINT0 = arrDCREINT0(arrDCREINT0_Index)
        newDCREINT0 = oldDCREINT0
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
fgSelect.LeftCol = 0

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





Private Sub txtSelect_AMJ_Change()
fgSelect.Clear
End Sub

Public Sub fgSelect_ForeColor(lColor As Long)
For I = 0 To fgSelect_arrIndex
  fgSelect.Col = I: fgSelect.CellForeColor = lColor
Next I

End Sub







Private Sub txtDetail_Change()
txtDetail_blnUpdate = True
cmdDetail_Control.Visible = True
End Sub

Private Sub txtDetail_KeyPress(KeyAscii As Integer)
Select Case txtDetail_Type
    Case "A": KeyAscii = convUCase(KeyAscii)
    Case "N": Call num_KeyAscii(KeyAscii)
    Case "C": Call num_KeyAsciiD(KeyAscii, txtDetail)
End Select
End Sub


Private Sub txtSelect_CREINTCLI_Change()
cmdSelect_Reset

End Sub

Private Sub txtSelect_CREINTCLI_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtSelect_CREINTDOS_Change()
cmdSelect_Reset

End Sub

Private Sub txtSelect_CREINTDOS_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub


Private Sub txtSelect_CREINTPER_Change()
cmdSelect_Reset

End Sub


Private Sub txtSelect_CREINTDEV_Change()
cmdSelect_Reset

End Sub

Private Sub txtSelect_CREINTDEV_Click()
cmdSelect_Reset

End Sub

Private Sub txtSelect_CREINTDEV_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtSelect_CREINTNAT_Change()
cmdSelect_Reset

End Sub

Private Sub txtSelect_CREINTNAT_Click()
cmdSelect_Reset

End Sub

Private Sub txtSelect_CREINTNAT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub






Public Sub fraDetail_Control()
Dim newX As String, oldX As String, X As String
Dim txtDetail_A As String, txtDetail_N As Long, txtDetail_D As Double, txtDetail_C As Currency
txtDetail.Visible = False
fgDetail.Row = txtDetail_Row
newX = Trim(txtDetail)
Select Case txtDetail_Type
    Case "A": txtDetail_A = newX
    Case "N": txtDetail_N = Val(newX)
    Case "D": txtDetail_D = num_CDec(newX)
    Case "C": txtDetail_C = num_CDec(newX)
End Select

fgDetail.Col = 4
    Select Case txtDetail_Field
       Case "CREINTT01": xDCREINT0.CREINTT01 = txtDetail_N
                         If oldDCREINT0.CREINTT01 = xDCREINT0.CREINTT01 Then
                            newDCREINT0.CREINTT01 = oldDCREINT0.CREINTT01
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT01 = xDCREINT0.CREINTT01
                            fgDetail = Format$(newDCREINT0.CREINTT01, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTT02": xDCREINT0.CREINTT02 = txtDetail_N
                         If oldDCREINT0.CREINTT02 = xDCREINT0.CREINTT02 Then
                            newDCREINT0.CREINTT02 = oldDCREINT0.CREINTT02
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT02 = xDCREINT0.CREINTT02
                            fgDetail = Format$(newDCREINT0.CREINTT02, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTT03": xDCREINT0.CREINTT03 = txtDetail_N
                         If oldDCREINT0.CREINTT03 = xDCREINT0.CREINTT03 Then
                            newDCREINT0.CREINTT03 = oldDCREINT0.CREINTT03
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT03 = xDCREINT0.CREINTT03
                            fgDetail = Format$(newDCREINT0.CREINTT03, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTT04": xDCREINT0.CREINTT04 = txtDetail_N
                         If oldDCREINT0.CREINTT04 = xDCREINT0.CREINTT04 Then
                            newDCREINT0.CREINTT04 = oldDCREINT0.CREINTT04
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT04 = xDCREINT0.CREINTT04
                            fgDetail = Format$(newDCREINT0.CREINTT04, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
                        
       Case "CREINTT11": xDCREINT0.CREINTT11 = txtDetail_N
                         If oldDCREINT0.CREINTT11 = xDCREINT0.CREINTT11 Then
                            newDCREINT0.CREINTT11 = oldDCREINT0.CREINTT11
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT11 = xDCREINT0.CREINTT11
                            fgDetail = Format$(newDCREINT0.CREINTT11, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTT12": xDCREINT0.CREINTT12 = txtDetail_N
                         If oldDCREINT0.CREINTT12 = xDCREINT0.CREINTT12 Then
                            newDCREINT0.CREINTT12 = oldDCREINT0.CREINTT12
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT12 = xDCREINT0.CREINTT12
                            fgDetail = Format$(newDCREINT0.CREINTT12, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTT13": xDCREINT0.CREINTT13 = txtDetail_N
                         If oldDCREINT0.CREINTT13 = xDCREINT0.CREINTT13 Then
                            newDCREINT0.CREINTT13 = oldDCREINT0.CREINTT13
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT13 = xDCREINT0.CREINTT13
                            fgDetail = Format$(newDCREINT0.CREINTT13, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTT14": xDCREINT0.CREINTT14 = txtDetail_N
                         If oldDCREINT0.CREINTT14 = xDCREINT0.CREINTT14 Then
                            newDCREINT0.CREINTT14 = oldDCREINT0.CREINTT14
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT14 = xDCREINT0.CREINTT14
                            fgDetail = Format$(newDCREINT0.CREINTT14, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If

       Case "CREINTT21": xDCREINT0.CREINTT21 = txtDetail_N
                         If oldDCREINT0.CREINTT21 = xDCREINT0.CREINTT21 Then
                            newDCREINT0.CREINTT21 = oldDCREINT0.CREINTT21
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT21 = xDCREINT0.CREINTT21
                            fgDetail = Format$(newDCREINT0.CREINTT21, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTT22": xDCREINT0.CREINTT22 = txtDetail_N
                         If oldDCREINT0.CREINTT22 = xDCREINT0.CREINTT22 Then
                            newDCREINT0.CREINTT22 = oldDCREINT0.CREINTT22
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT22 = xDCREINT0.CREINTT22
                            fgDetail = Format$(newDCREINT0.CREINTT22, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTT23": xDCREINT0.CREINTT23 = txtDetail_N
                         If oldDCREINT0.CREINTT23 = xDCREINT0.CREINTT23 Then
                            newDCREINT0.CREINTT23 = oldDCREINT0.CREINTT23
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT23 = xDCREINT0.CREINTT23
                            fgDetail = Format$(newDCREINT0.CREINTT23, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTT24": xDCREINT0.CREINTT24 = txtDetail_N
                         If oldDCREINT0.CREINTT24 = xDCREINT0.CREINTT24 Then
                            newDCREINT0.CREINTT24 = oldDCREINT0.CREINTT24
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT24 = xDCREINT0.CREINTT24
                            fgDetail = Format$(newDCREINT0.CREINTT24, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If

       Case "CREINTT31": xDCREINT0.CREINTT31 = txtDetail_N
                         If oldDCREINT0.CREINTT31 = xDCREINT0.CREINTT31 Then
                            newDCREINT0.CREINTT31 = oldDCREINT0.CREINTT31
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT31 = xDCREINT0.CREINTT31
                            fgDetail = Format$(newDCREINT0.CREINTT31, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTT32": xDCREINT0.CREINTT32 = txtDetail_N
                         If oldDCREINT0.CREINTT32 = xDCREINT0.CREINTT32 Then
                            newDCREINT0.CREINTT32 = oldDCREINT0.CREINTT32
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT32 = xDCREINT0.CREINTT32
                            fgDetail = Format$(newDCREINT0.CREINTT32, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTT33": xDCREINT0.CREINTT33 = txtDetail_N
                         If oldDCREINT0.CREINTT33 = xDCREINT0.CREINTT33 Then
                            newDCREINT0.CREINTT33 = oldDCREINT0.CREINTT33
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT33 = xDCREINT0.CREINTT33
                            fgDetail = Format$(newDCREINT0.CREINTT33, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTT34": xDCREINT0.CREINTT34 = txtDetail_N
                         If oldDCREINT0.CREINTT34 = xDCREINT0.CREINTT34 Then
                            newDCREINT0.CREINTT34 = oldDCREINT0.CREINTT34
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT34 = xDCREINT0.CREINTT34
                            fgDetail = Format$(newDCREINT0.CREINTT34, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If

       Case "CREINTT41": xDCREINT0.CREINTT41 = txtDetail_N
                         If oldDCREINT0.CREINTT41 = xDCREINT0.CREINTT41 Then
                            newDCREINT0.CREINTT41 = oldDCREINT0.CREINTT41
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT41 = xDCREINT0.CREINTT41
                            fgDetail = Format$(newDCREINT0.CREINTT41, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTT42": xDCREINT0.CREINTT42 = txtDetail_N
                         If oldDCREINT0.CREINTT42 = xDCREINT0.CREINTT42 Then
                            newDCREINT0.CREINTT42 = oldDCREINT0.CREINTT42
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT42 = xDCREINT0.CREINTT42
                            fgDetail = Format$(newDCREINT0.CREINTT42, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTT43": xDCREINT0.CREINTT43 = txtDetail_N
                         If oldDCREINT0.CREINTT43 = xDCREINT0.CREINTT43 Then
                            newDCREINT0.CREINTT43 = oldDCREINT0.CREINTT43
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT43 = xDCREINT0.CREINTT43
                            fgDetail = Format$(newDCREINT0.CREINTT43, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTT44": xDCREINT0.CREINTT44 = txtDetail_N
                         If oldDCREINT0.CREINTT44 = xDCREINT0.CREINTT44 Then
                            newDCREINT0.CREINTT44 = oldDCREINT0.CREINTT44
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTT44 = xDCREINT0.CREINTT44
                            fgDetail = Format$(newDCREINT0.CREINTT44, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
'_________________________________________________________________________________________________
       Case "CREINTM01": xDCREINT0.CREINTM01 = txtDetail_N
                         If oldDCREINT0.CREINTM01 = xDCREINT0.CREINTM01 Then
                            newDCREINT0.CREINTM01 = oldDCREINT0.CREINTM01
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM01 = xDCREINT0.CREINTM01
                            fgDetail = Format$(newDCREINT0.CREINTM01, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTM02": xDCREINT0.CREINTM02 = txtDetail_N
                         If oldDCREINT0.CREINTM02 = xDCREINT0.CREINTM02 Then
                            newDCREINT0.CREINTM02 = oldDCREINT0.CREINTM02
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM02 = xDCREINT0.CREINTM02
                            fgDetail = Format$(newDCREINT0.CREINTM02, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTM03": xDCREINT0.CREINTM03 = txtDetail_N
                         If oldDCREINT0.CREINTM03 = xDCREINT0.CREINTM03 Then
                            newDCREINT0.CREINTM03 = oldDCREINT0.CREINTM03
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM03 = xDCREINT0.CREINTM03
                            fgDetail = Format$(newDCREINT0.CREINTM03, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTM04": xDCREINT0.CREINTM04 = txtDetail_N
                         If oldDCREINT0.CREINTM04 = xDCREINT0.CREINTM04 Then
                            newDCREINT0.CREINTM04 = oldDCREINT0.CREINTM04
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM04 = xDCREINT0.CREINTM04
                            fgDetail = Format$(newDCREINT0.CREINTM04, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
                        
       Case "CREINTM11": xDCREINT0.CREINTM11 = txtDetail_N
                         If oldDCREINT0.CREINTM11 = xDCREINT0.CREINTM11 Then
                            newDCREINT0.CREINTM11 = oldDCREINT0.CREINTM11
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM11 = xDCREINT0.CREINTM11
                            fgDetail = Format$(newDCREINT0.CREINTM11, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTM12": xDCREINT0.CREINTM12 = txtDetail_N
                         If oldDCREINT0.CREINTM12 = xDCREINT0.CREINTM12 Then
                            newDCREINT0.CREINTM12 = oldDCREINT0.CREINTM12
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM12 = xDCREINT0.CREINTM12
                            fgDetail = Format$(newDCREINT0.CREINTM12, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTM13": xDCREINT0.CREINTM13 = txtDetail_N
                         If oldDCREINT0.CREINTM13 = xDCREINT0.CREINTM13 Then
                            newDCREINT0.CREINTM13 = oldDCREINT0.CREINTM13
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM13 = xDCREINT0.CREINTM13
                            fgDetail = Format$(newDCREINT0.CREINTM13, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTM14": xDCREINT0.CREINTM14 = txtDetail_N
                         If oldDCREINT0.CREINTM14 = xDCREINT0.CREINTM14 Then
                            newDCREINT0.CREINTM14 = oldDCREINT0.CREINTM14
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM14 = xDCREINT0.CREINTM14
                            fgDetail = Format$(newDCREINT0.CREINTM14, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If

       Case "CREINTM21": xDCREINT0.CREINTM21 = txtDetail_N
                         If oldDCREINT0.CREINTM21 = xDCREINT0.CREINTM21 Then
                            newDCREINT0.CREINTM21 = oldDCREINT0.CREINTM21
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM21 = xDCREINT0.CREINTM21
                            fgDetail = Format$(newDCREINT0.CREINTM21, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTM22": xDCREINT0.CREINTM22 = txtDetail_N
                         If oldDCREINT0.CREINTM22 = xDCREINT0.CREINTM22 Then
                            newDCREINT0.CREINTM22 = oldDCREINT0.CREINTM22
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM22 = xDCREINT0.CREINTM22
                            fgDetail = Format$(newDCREINT0.CREINTM22, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTM23": xDCREINT0.CREINTM23 = txtDetail_N
                         If oldDCREINT0.CREINTM23 = xDCREINT0.CREINTM23 Then
                            newDCREINT0.CREINTM23 = oldDCREINT0.CREINTM23
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM23 = xDCREINT0.CREINTM23
                            fgDetail = Format$(newDCREINT0.CREINTM23, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTM24": xDCREINT0.CREINTM24 = txtDetail_N
                         If oldDCREINT0.CREINTM24 = xDCREINT0.CREINTM24 Then
                            newDCREINT0.CREINTM24 = oldDCREINT0.CREINTM24
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM24 = xDCREINT0.CREINTM24
                            fgDetail = Format$(newDCREINT0.CREINTM24, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If

       Case "CREINTM31": xDCREINT0.CREINTM31 = txtDetail_N
                         If oldDCREINT0.CREINTM31 = xDCREINT0.CREINTM31 Then
                            newDCREINT0.CREINTM31 = oldDCREINT0.CREINTM31
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM31 = xDCREINT0.CREINTM31
                            fgDetail = Format$(newDCREINT0.CREINTM31, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTM32": xDCREINT0.CREINTM32 = txtDetail_N
                         If oldDCREINT0.CREINTM32 = xDCREINT0.CREINTM32 Then
                            newDCREINT0.CREINTM32 = oldDCREINT0.CREINTM32
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM32 = xDCREINT0.CREINTM32
                            fgDetail = Format$(newDCREINT0.CREINTM32, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTM33": xDCREINT0.CREINTM33 = txtDetail_N
                         If oldDCREINT0.CREINTM33 = xDCREINT0.CREINTM33 Then
                            newDCREINT0.CREINTM33 = oldDCREINT0.CREINTM33
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM33 = xDCREINT0.CREINTM33
                            fgDetail = Format$(newDCREINT0.CREINTM33, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTM34": xDCREINT0.CREINTM34 = txtDetail_N
                         If oldDCREINT0.CREINTM34 = xDCREINT0.CREINTM34 Then
                            newDCREINT0.CREINTM34 = oldDCREINT0.CREINTM34
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM34 = xDCREINT0.CREINTM34
                            fgDetail = Format$(newDCREINT0.CREINTM34, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If

       Case "CREINTM41": xDCREINT0.CREINTM41 = txtDetail_N
                         If oldDCREINT0.CREINTM41 = xDCREINT0.CREINTM41 Then
                            newDCREINT0.CREINTM41 = oldDCREINT0.CREINTM41
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM41 = xDCREINT0.CREINTM41
                            fgDetail = Format$(newDCREINT0.CREINTM41, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTM42": xDCREINT0.CREINTM42 = txtDetail_N
                         If oldDCREINT0.CREINTM42 = xDCREINT0.CREINTM42 Then
                            newDCREINT0.CREINTM42 = oldDCREINT0.CREINTM42
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM42 = xDCREINT0.CREINTM42
                            fgDetail = Format$(newDCREINT0.CREINTM42, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTM43": xDCREINT0.CREINTM43 = txtDetail_N
                         If oldDCREINT0.CREINTM43 = xDCREINT0.CREINTM43 Then
                            newDCREINT0.CREINTM43 = oldDCREINT0.CREINTM43
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM43 = xDCREINT0.CREINTM43
                            fgDetail = Format$(newDCREINT0.CREINTM43, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "CREINTM44": xDCREINT0.CREINTM44 = txtDetail_N
                         If oldDCREINT0.CREINTM44 = xDCREINT0.CREINTM44 Then
                            newDCREINT0.CREINTM44 = oldDCREINT0.CREINTM44
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDCREINT0.CREINTM44 = xDCREINT0.CREINTM44
                            fgDetail = Format$(newDCREINT0.CREINTM44, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If

End Select
fgDetail.Row = fgDetail_RowClick
txtDetail_blnUpdate = False: cmdDetail_Control.Visible = False
cmdDetail_Update.Visible = DCREINT0_Aut.Saisir
End Sub



Public Sub cmdSelect_SQL_Import_TA()
Dim wFile As String, wAAAA As Long
Dim K As Long, I As Long

Call lstErr_Clear(lstErr, cmdContext, "> mise TA => DCREINT0"): DoEvents

Call DTPicker_Control(txtSelect_CREINTPER, wAmjMax)
wAAAA = Val(Mid$(wAmjMax, 1, 4))

arrTRIM(0).CRETADEB = "00000000": arrTRIM(0).CRETAFIN = wAAAA - 1 & "1231"

arrTRIM(1).CRETADEB = arrTRIM(0).CRETAFIN: arrTRIM(1).CRETAFIN = wAAAA & "0331"
arrTRIM(2).CRETADEB = arrTRIM(1).CRETAFIN: arrTRIM(2).CRETAFIN = wAAAA & "0630"
arrTRIM(3).CRETADEB = arrTRIM(2).CRETAFIN: arrTRIM(3).CRETAFIN = wAAAA & "0930"
arrTRIM(4).CRETADEB = arrTRIM(3).CRETAFIN: arrTRIM(4).CRETAFIN = wAAAA & "1231"
arrTRIM(5).CRETADEB = arrTRIM(4).CRETAFIN: arrTRIM(5).CRETAFIN = wAAAA + 1 & "0331"
arrTRIM(6).CRETADEB = arrTRIM(5).CRETAFIN: arrTRIM(6).CRETAFIN = wAAAA + 1 & "0630"
arrTRIM(7).CRETADEB = arrTRIM(6).CRETAFIN: arrTRIM(7).CRETAFIN = wAAAA + 1 & "0930"
arrTRIM(7).CRETAMIN = 0: arrTRIM(7).CRETATAU = 0: arrTRIM(7).CRETAMARGE = 0
arrTRIM(8).CRETADEB = arrTRIM(7).CRETAFIN: arrTRIM(8).CRETAFIN = wAAAA + 1 & "1231"
arrTRIM(8).CRETAMIN = 0: arrTRIM(8).CRETATAU = 0: arrTRIM(8).CRETAMARGE = 0

arrTRIM(9).CRETADEB = arrTRIM(8).CRETAFIN: arrTRIM(9).CRETAFIN = wAAAA + 2 & "0331"
arrTRIM(10).CRETADEB = arrTRIM(9).CRETAFIN: arrTRIM(10).CRETAFIN = wAAAA + 2 & "0630"
arrTRIM(11).CRETADEB = arrTRIM(10).CRETAFIN: arrTRIM(11).CRETAFIN = wAAAA + 2 & "0930"
arrTRIM(12).CRETADEB = arrTRIM(11).CRETAFIN: arrTRIM(12).CRETAFIN = wAAAA + 2 & "1231"
arrTRIM(13).CRETADEB = arrTRIM(12).CRETAFIN: arrTRIM(13).CRETAFIN = wAAAA + 3 & "0331"
arrTRIM(14).CRETADEB = arrTRIM(13).CRETAFIN: arrTRIM(14).CRETAFIN = wAAAA + 3 & "0630"
arrTRIM(15).CRETADEB = arrTRIM(14).CRETAFIN: arrTRIM(15).CRETAFIN = wAAAA + 3 & "0930"
arrTRIM(16).CRETADEB = arrTRIM(15).CRETAFIN: arrTRIM(16).CRETAFIN = wAAAA + 3 & "1231"
arrTRIM(17).CRETADEB = arrTRIM(16).CRETAFIN: arrTRIM(17).CRETAFIN = wAAAA + 4 & "0331"
arrTRIM(18).CRETADEB = arrTRIM(17).CRETAFIN: arrTRIM(18).CRETAFIN = wAAAA + 4 & "0630"
arrTRIM(19).CRETADEB = arrTRIM(18).CRETAFIN: arrTRIM(19).CRETAFIN = wAAAA + 4 & "0930"
arrTRIM(20).CRETADEB = arrTRIM(19).CRETAFIN: arrTRIM(20).CRETAFIN = wAAAA + 4 & "1231"
arrTRIM(21).CRETADEB = arrTRIM(20).CRETAFIN: arrTRIM(21).CRETAFIN = "99999999"

For K = 0 To lstW.ListCount - 1
    lstW.ListIndex = K
    wFile = paramEditionSplf_Folder & "\" & wFolder_CREGS601P1 & "\" & lstW.Text
    For I = 0 To 21
        arrTRIM(I).CRETAMIN = 0: arrTRIM(I).CRETATAU = 0: arrTRIM(I).CRETAMARGE = 0
    
    Next I
    Call cmdSelect_SQL_Import_TA_Detail(wFile)
Next K
End Sub

Public Sub cmdSelect_SQL_Import_TA_Detail(lFile As String)
Dim xIn As String, blnCREINTDOS As Boolean, blnErr As Boolean, blnExit As Boolean
Dim K As Long, K1 As Long, K2 As Long
Dim xWhere As String, xSql As String
Dim blnCREINTECH As Boolean
Dim wAMJ As String, X As String
Dim wCRETADEB As String, wCRETAFIN As String, wCRETAMIN As Currency, wCRETATAU As Double
Dim blnCREEVE_Ok As Boolean, blnEchu As Boolean, blnEch_Ok As Boolean
Dim wNBJ As Long, wNBJ1 As Long
Dim wCur As Currency
Dim wCREINTUAMJ As Long, wCREINTUHMS As Long
Dim kX1 As Integer
On Error GoTo Exit_Sub

arrDCRETA_Nb = 0
blnCREINTDOS = False
Open lFile For Input As #3
Line Input #3, xIn
If IsNumeric(Mid$(xIn, 5, 8)) Then
    wCREINTUAMJ = Mid$(xIn, 5, 8)
Else
    wCREINTUAMJ = DSys
End If
If IsNumeric(Mid$(xIn, 64, 6)) Then
    wCREINTUHMS = Mid$(xIn, 64, 6)
Else
    wCREINTUHMS = Time
End If

Do Until EOF(3)
    Line Input #3, xIn
    xIn = Trim(xIn)
'_________________________________________________________________________
If Not blnCREINTDOS Then
        K = InStr(xIn, "Dossier")
        If K > 0 Then
            blnCREINTDOS = True
            For K1 = K + 7 To K + 20
                If Mid$(xIn, K1, 1) <> " " Then Exit For
            Next K1
            oldDCREINT0.CREINTDOS = Val(Mid$(xIn, K1, 7))
            oldDCREINT0.CREINTPRE = Val(Mid$(xIn, K1 + 9, 2))
            xWhere = " where CREINTPER = " & wAmjMax & " and CREINTDOS = " & oldDCREINT0.CREINTDOS & " and CREINTPRE = " & oldDCREINT0.CREINTPRE
            arrDCREINT0_SQL xWhere
            If arrDCREINT0_Nb = 1 Then
                blnErr = False
                oldDCREINT0 = arrDCREINT0(1)
                newDCREINT0 = oldDCREINT0
                Call lstErr_ChangeLastItem(lstErr, cmdContext, "- " & xIn): DoEvents
            Else
                 blnErr = True
                 Exit Do
            End If
        End If
'_________________________________________________________________________
Else
    If Not blnCREINTECH Then
        K = InStr(xIn, "Echéance")
        If K > 0 Then blnCREINTECH = True
    Else
        kX1 = InStr(xIn, "|")
        If IsNumeric(Mid$(xIn, kX1 + 7, 2)) Then
            X = Replace(Mid$(xIn, kX1 + 7, 8), " ", "0")
            arrDCRETA_Nb = arrDCRETA_Nb + 1
            arrDCRETA(arrDCRETA_Nb).CRETADEB = ""
            arrDCRETA(arrDCRETA_Nb).CRETAFIN = ""
            arrDCRETA(arrDCRETA_Nb).CRETAMARGE = 0
            arrDCRETA(arrDCRETA_Nb).CRETATAU = 0
            arrDCRETA(arrDCRETA_Nb).CRETAMIN = 0
            Call dateJMA6_AMJ(X, wAMJ)
            arrDCRETA(arrDCRETA_Nb).CRETAFIN = "20" & wAMJ
            X = Trim(Replace(Mid$(xIn, kX1 + 59, 13), ".", ""))
            If X <> "" Then arrDCRETA(arrDCRETA_Nb).CRETAMIN = CCur(X)
            
        Else
            
            If InStr(Mid$(xIn, 1, 20), "Total") > 0 Then Exit Do
        End If
    End If
End If

Loop

Close #3

'_________________________________________________________________________

If Not blnErr Then
    xSql = "select * from " & paramIBM_Library_SAB & ".ZCREEVE0 " _
         & " where CREEVEDOS =" & oldDCREINT0.CREINTDOS _
         & " and   CREEVEPRE =" & oldDCREINT0.CREINTPRE _
         & " and   CREEVETYP in ('02','03') And CREEVEREG > 0 " _
         & " order by CREEVEDEB "
    Set rsSab = cnsab.Execute(xSql)
    
    If rsSab.EOF Then
        blnEch_Ok = False
        xSql = "select * from " & paramIBM_Library_SAB & ".ZCREEVE0 " _
             & " where CREEVEDOS =" & oldDCREINT0.CREINTDOS _
             & " and   CREEVEPRE =" & oldDCREINT0.CREINTPRE _
             & " and   CREEVETYP in ('02','03') " _
             & " order by CREEVEDEB "
        Set rsSab = cnsab.Execute(xSql)
        If rsSab.EOF Then
            MsgBox "? Pas de données dans ZCREEVE0", vbCritical, "Dossier : " & oldDCREINT0.CREINTDOS & " - " & oldDCREINT0.CREINTPRE
            Exit Sub
        End If
        'Call lstErr_AddItem(lstErr, cmdContext, "? PAS de données dans ZCREEVE0"): DoEvents
        'Exit Sub
    Else
        blnEch_Ok = True
    End If
    
    wCRETAFIN = rsSab("CREEVEFIN") + 19000000
    If wCRETAFIN = arrDCRETA(1).CRETAFIN Then
        blnEchu = True
    Else
        blnEchu = False
    End If


    Do While Not rsSab.EOF
        wCRETADEB = rsSab("CREEVEDEB") + 19000000
        wCRETAFIN = rsSab("CREEVEFIN") + 19000000
        wCRETAMIN = rsSab("CREEVEMIN")
        wCRETATAU = rsSab("CREEVETAU")
        blnCREEVE_Ok = False
        
        For K = 1 To arrDCRETA_Nb
            If blnEchu Then
                If wCRETAFIN = arrDCRETA(K).CRETAFIN Then
                    If Not blnEch_Ok Then wCRETAMIN = arrDCRETA(K).CRETAMIN
                    If wCRETAMIN = arrDCRETA(K).CRETAMIN Then
                        blnCREEVE_Ok = True
                        arrDCRETA(K).CRETADEB = wCRETADEB
                        arrDCRETA(K).CRETATAU = wCRETATAU
                        Exit For
                    Else
                        'blnErr = True
                        MsgBox "Période " & wCRETADEB & " - " & wCRETAFIN & vbCrLf _
                        & "Mt intérêts  ZCREVE0 : " & wCRETAMIN & " <> TA : " & arrDCRETA(K).CRETAMIN _
                        , vbCritical, "Dossier : " & oldDCREINT0.CREINTDOS & " - " & oldDCREINT0.CREINTPRE
                    End If
                End If
            Else
                  If wCRETADEB = arrDCRETA(K).CRETAFIN Then
                     If Not blnEch_Ok Then wCRETAMIN = arrDCRETA(K).CRETAMIN
                     If wCRETAMIN = arrDCRETA(K).CRETAMIN Then
                         blnCREEVE_Ok = True
                         arrDCRETA(K).CRETADEB = wCRETADEB
                         arrDCRETA(K).CRETAFIN = wCRETAFIN
                         arrDCRETA(K).CRETATAU = wCRETATAU
                         Exit For
                     Else
                        MsgBox "Période " & wCRETADEB & " - " & wCRETAFIN & vbCrLf _
                        & "Mt intérêts  ZCREVE0 : " & wCRETAMIN & " <> TA : " & arrDCRETA(K).CRETAMIN _
                        , vbCritical, "Dossier : " & oldDCREINT0.CREINTDOS & " - " & oldDCREINT0.CREINTPRE
                         'blnErr = True
                         'MsgBox "wCRETAMIN <> " & wCRETADEB & " " & wCRETAFIN, vbCritical, "Dossier : " & oldDCREINT0.CREINTDOS & " - " & oldDCREINT0.CREINTPRE
                     End If
                End If
            End If
        Next K
        rsSab.MoveNext
    Loop

'_________________________________________________________________________

    If Not blnErr Then
       If arrDCRETA(1).CRETADEB = "" Then
           xSql = "select CREEVEREG from " & paramIBM_Library_SAB & ".ZCREEVE0 " _
            & " where CREEVEDOS =" & oldDCREINT0.CREINTDOS _
            & " and   CREEVEPRE =" & oldDCREINT0.CREINTPRE _
            & " and   CREEVETYP =  '00' "
            Set rsSab = cnsab.Execute(xSql)
            
            If Not rsSab.EOF Then
                arrDCRETA(1).CRETADEB = rsSab("CREEVEREG") + 19000000
            Else
                arrDCRETA(1).CRETADEB = arrDCRETA(1).CRETAFIN
            End If

       End If
       
       For K = 1 To arrDCRETA_Nb
            
            If arrDCRETA(K).CRETADEB = "" Then arrDCRETA(K).CRETADEB = arrDCRETA(K - 1).CRETAFIN
            wNBJ = DateDiff("d", Format(arrDCRETA(K).CRETADEB, "@@@@/@@/@@"), Format(arrDCRETA(K).CRETAFIN, "@@@@/@@/@@"))
            blnExit = False
            For K1 = 0 To 21
                If arrDCRETA(K).CRETADEB <= arrTRIM(K1).CRETAFIN Then
                    If arrDCRETA(K).CRETAFIN <= arrTRIM(K1).CRETAFIN Then blnExit = True
                    
                    wCRETADEB = IIf(arrDCRETA(K).CRETADEB > arrTRIM(K1).CRETADEB, arrDCRETA(K).CRETADEB, arrTRIM(K1).CRETADEB)
                    wCRETAFIN = IIf(arrDCRETA(K).CRETAFIN < arrTRIM(K1).CRETAFIN, arrDCRETA(K).CRETAFIN, arrTRIM(K1).CRETAFIN)
                    wNBJ1 = DateDiff("d", Format(wCRETADEB, "@@@@/@@/@@"), Format(wCRETAFIN, "@@@@/@@/@@"))
                    'particularité SAB :
                    If wCRETADEB = arrDCRETA(1).CRETADEB Then wNBJ1 = wNBJ1 + 1
                    If wCRETAFIN = arrDCRETA(arrDCRETA_Nb).CRETAFIN Then wNBJ1 = wNBJ1 - 1
                    
                    If wNBJ <> 0 Then
                        wCur = arrDCRETA(K).CRETAMIN * wNBJ1 / wNBJ
                        arrTRIM(K1).CRETAMIN = arrTRIM(K1).CRETAMIN + wCur
                        If arrDCRETA(K).CRETATAU = 0 Then arrDCRETA(K).CRETATAU = oldDCREINT0.CREINTTOF + oldDCREINT0.CREINTTOM
                        If oldDCREINT0.CREINTTOM > 0 And arrDCRETA(K).CRETATAU > 0 Then
                            arrTRIM(K1).CRETAMARGE = arrTRIM(K1).CRETAMARGE + wCur * oldDCREINT0.CREINTTOM / arrDCRETA(K).CRETATAU
                        End If
                    End If
                End If
                If blnExit Then Exit For
            Next K1
        Next K
        
        newDCREINT0.CREINTT01 = arrTRIM(1).CRETAMIN: newDCREINT0.CREINTM01 = arrTRIM(1).CRETAMARGE
        newDCREINT0.CREINTT02 = arrTRIM(2).CRETAMIN: newDCREINT0.CREINTM02 = arrTRIM(2).CRETAMARGE
        newDCREINT0.CREINTT03 = arrTRIM(3).CRETAMIN: newDCREINT0.CREINTM03 = arrTRIM(3).CRETAMARGE
        newDCREINT0.CREINTT04 = arrTRIM(4).CRETAMIN: newDCREINT0.CREINTM04 = arrTRIM(4).CRETAMARGE
        newDCREINT0.CREINTT11 = arrTRIM(5).CRETAMIN: newDCREINT0.CREINTM11 = arrTRIM(5).CRETAMARGE
        newDCREINT0.CREINTT12 = arrTRIM(6).CRETAMIN: newDCREINT0.CREINTM12 = arrTRIM(6).CRETAMARGE
        newDCREINT0.CREINTT13 = arrTRIM(7).CRETAMIN: newDCREINT0.CREINTM13 = arrTRIM(7).CRETAMARGE
        newDCREINT0.CREINTT14 = arrTRIM(8).CRETAMIN: newDCREINT0.CREINTM14 = arrTRIM(8).CRETAMARGE
        newDCREINT0.CREINTT21 = arrTRIM(9).CRETAMIN: newDCREINT0.CREINTM21 = arrTRIM(9).CRETAMARGE
        newDCREINT0.CREINTT22 = arrTRIM(10).CRETAMIN: newDCREINT0.CREINTM22 = arrTRIM(10).CRETAMARGE
        newDCREINT0.CREINTT23 = arrTRIM(11).CRETAMIN: newDCREINT0.CREINTM23 = arrTRIM(11).CRETAMARGE
        newDCREINT0.CREINTT24 = arrTRIM(12).CRETAMIN: newDCREINT0.CREINTM24 = arrTRIM(12).CRETAMARGE
        newDCREINT0.CREINTT31 = arrTRIM(13).CRETAMIN: newDCREINT0.CREINTM31 = arrTRIM(13).CRETAMARGE
        newDCREINT0.CREINTT32 = arrTRIM(14).CRETAMIN: newDCREINT0.CREINTM32 = arrTRIM(14).CRETAMARGE
        newDCREINT0.CREINTT33 = arrTRIM(15).CRETAMIN: newDCREINT0.CREINTM33 = arrTRIM(15).CRETAMARGE
        newDCREINT0.CREINTT34 = arrTRIM(16).CRETAMIN: newDCREINT0.CREINTM34 = arrTRIM(16).CRETAMARGE
        newDCREINT0.CREINTT41 = arrTRIM(17).CRETAMIN: newDCREINT0.CREINTM41 = arrTRIM(17).CRETAMARGE
        newDCREINT0.CREINTT42 = arrTRIM(18).CRETAMIN: newDCREINT0.CREINTM42 = arrTRIM(18).CRETAMARGE
        newDCREINT0.CREINTT43 = arrTRIM(19).CRETAMIN: newDCREINT0.CREINTM43 = arrTRIM(19).CRETAMARGE
        newDCREINT0.CREINTT44 = arrTRIM(20).CRETAMIN: newDCREINT0.CREINTM44 = arrTRIM(20).CRETAMARGE
        
        newDCREINT0.CREINTUAMJ = wCREINTUAMJ
        newDCREINT0.CREINTUHMS = wCREINTUHMS
        
        cmdDCREINT0_Update
        Call lstErr_AddItem(lstErr, cmdContext, "- mise à jour du fichier DCREINT0"): DoEvents
    End If
End If
Exit Sub

Exit_Sub:
Close
MsgBox lFile & vbCrLf & Error, vbCritical, "Dossier : " & oldDCREINT0.CREINTDOS & " - " & oldDCREINT0.CREINTPRE
End Sub

Public Sub cmdSelect_SQL_Import_TA_Init()
Dim objFolder, objFiles_Open
Dim fsoFile As File, fsoFile2 As File
Dim currentFileName As String

wFolder_CREGS601P1 = InputBox("Préciser le répertoire :" & vbCrLf & vbCrLf & "'Production' ou 'Corbeille'", "DCREINT : Répertoire des états CREGS601P1", paramEnvironnement)
If Trim(wFolder_CREGS601P1) = "" Then Exit Sub

Me.Enabled = False: Me.MousePointer = vbHourglass
lstW.Visible = False
lstW.Clear
Set objFolder = msFileSystem.GetFolder(paramEditionSplf_Folder & "\" & wFolder_CREGS601P1)
Set objFiles_Open = objFolder.Files
For Each fsoFile In objFiles_Open
    currentFileName = fsoFile.Name
    If InStr(currentFileName, usrName_UCase) > 0 Then
         If InStr(currentFileName, "CREGS601P1") > 0 Then lstW.AddItem currentFileName
    End If
Next fsoFile
lstW.Visible = True
Me.Enabled = True: Me.MousePointer = 0

End Sub
