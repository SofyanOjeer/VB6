VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDGAPPIS0 
   AutoRedraw      =   -1  'True
   Caption         =   "DGAPPIS0 : maintenance"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13575
   Icon            =   "DGAPPIS0.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9315
   ScaleWidth      =   13575
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   0
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8685
      Left            =   -45
      TabIndex        =   2
      Top             =   480
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   15319
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "DGAPPIS0"
      TabPicture(0)   =   "DGAPPIS0.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "......."
      TabPicture(1)   =   "DGAPPIS0.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDetail"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraDetail 
         BackColor       =   &H00A0E0FF&
         Height          =   6852
         Left            =   -71640
         TabIndex        =   21
         Top             =   960
         Width           =   7500
         Begin VB.CommandButton cmdDetail_Update 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Enregistrer"
            Height          =   500
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   6150
            Width           =   1095
         End
         Begin VB.CommandButton cmdDetail_Delete 
            BackColor       =   &H000000FF&
            Caption         =   "Supprimer"
            Height          =   500
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   6150
            Width           =   1095
         End
         Begin VB.CommandButton cmdDetail_Copy 
            BackColor       =   &H00FF80FF&
            Caption         =   "Copier"
            Height          =   500
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   6150
            Width           =   1095
         End
         Begin VB.CommandButton cmdDetail_Quit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Abandonner"
            Enabled         =   0   'False
            Height          =   500
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   6150
            Width           =   1095
         End
         Begin VB.CommandButton cmdDetail_Control 
            BackColor       =   &H0080C0FF&
            Caption         =   "Saisie"
            Height          =   500
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   6120
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   5865
            Left            =   135
            TabIndex        =   27
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
         Height          =   8232
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   13296
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
            Left            =   11520
            TabIndex        =   13
            Top             =   120
            Visible         =   0   'False
            Width           =   1452
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   555
            Left            =   11520
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   480
            Width           =   1335
         End
         Begin VB.Frame fraSelect_Options 
            BackColor       =   &H00F0FFFF&
            Height          =   1005
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   10995
            Begin VB.TextBox txtSelect_GAPPISRUB 
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
               Left            =   9465
               TabIndex        =   29
               Top             =   570
               Width           =   1212
            End
            Begin VB.TextBox txtSelect_GAPPISNUO 
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
               Left            =   7995
               TabIndex        =   28
               Top             =   615
               Width           =   1212
            End
            Begin VB.ComboBox txtSelect_GAPPISSEN 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   5265
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   540
               Width           =   852
            End
            Begin VB.ComboBox txtSelect_GAPPISDEV 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   2895
               Sorted          =   -1  'True
               TabIndex        =   16
               Text            =   "txtSelect_GAPPISDEV"
               Top             =   570
               Width           =   852
            End
            Begin VB.ComboBox txtSelect_GAPPISOPE 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   6360
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   525
               Width           =   1332
            End
            Begin VB.ComboBox txtSelect_GAPPISTAB 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   4065
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   555
               Width           =   852
            End
            Begin VB.TextBox txtSelect_DGAPPISCLI 
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
               Left            =   1575
               TabIndex        =   11
               Top             =   555
               Width           =   1212
            End
            Begin MSComCtl2.DTPicker txtSelect_DGAPPISPER 
               Height          =   300
               Left            =   75
               TabIndex        =   7
               Top             =   540
               Width           =   1335
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
               Format          =   57344003
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_GAPPISRUB 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Rubrique"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9795
               TabIndex        =   31
               Top             =   255
               Width           =   615
            End
            Begin VB.Label lblSelect_GAPPISNUO 
               BackColor       =   &H00F0FFFF&
               Caption         =   "N° opération"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   8115
               TabIndex        =   30
               Top             =   255
               Width           =   1005
            End
            Begin VB.Label lblSelect_GAPPISSEN 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Sens"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5385
               TabIndex        =   19
               Top             =   195
               Width           =   615
            End
            Begin VB.Label lblSelect_GAPPISDEV 
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
               Height          =   255
               Left            =   2910
               TabIndex        =   18
               Top             =   150
               Width           =   615
            End
            Begin VB.Label lblSelect_GAPPISOPE 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Opération"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6360
               TabIndex        =   17
               Top             =   210
               Width           =   855
            End
            Begin VB.Label lblSelect_GAPPISTAB 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Code état (1 2 3 )"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3945
               TabIndex        =   12
               Top             =   165
               Width           =   1215
            End
            Begin VB.Label lblSelect_DGAPPISCLI 
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
               Height          =   255
               Left            =   1785
               TabIndex        =   10
               Top             =   195
               Width           =   615
            End
            Begin VB.Label lblSelect_DGAPPISPER 
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
               Height          =   255
               Left            =   255
               TabIndex        =   9
               Top             =   195
               Width           =   615
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6828
            Left            =   240
            TabIndex        =   5
            Top             =   1320
            Width           =   11472
            _ExtentX        =   20241
            _ExtentY        =   12039
            _Version        =   393216
            Rows            =   1
            Cols            =   12
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
            FormatString    =   $"DGAPPIS0.frx":0342
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
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13080
      Picture         =   "DGAPPIS0.frx":0408
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
Attribute VB_Name = "frmDGAPPIS0"
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
Dim DGAPPIS0_Aut As typeAuthorization
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
Dim xDGAPPIS0 As typeDGAPPIS0, newDGAPPIS0 As typeDGAPPIS0, oldDGAPPIS0 As typeDGAPPIS0
Dim arrDGAPPIS0() As typeDGAPPIS0, arrDGAPPIS0_Nb As Long, arrDGAPPIS0_Max As Long, arrDGAPPIS0_Index As Long

Dim txtDetail_Type As String, txtDetail_Update As String
Dim txtDetail_Field As String
Dim txtDetail_blnUpdate As Boolean, txtDetail_ColorUpdate As Long
Dim txtDetail_Row As Integer

Dim arrDGAPPISVEC(8) As String, arrDGAPPISVEC_AMJ(8) As Long
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
cmdDetail_Delete.Visible = DGAPPIS0_Aut.Saisir
cmdDetail_Copy.Visible = DGAPPIS0_Aut.Saisir
cmdDetail_Update.Visible = False
cmdDetail_Quit.Enabled = True
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString

    
For I = 1 To 45
         
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



wField = Trim(Mid$(ddsDGAPPIS0(lIndex), 4, 10))
fgDetail.Col = 0: fgDetail.Text = Mid$(ddsDGAPPIS0(lIndex), 1, 2)
fgDetail.Col = 1: fgDetail.Text = wField
fgDetail.Col = 2: fgDetail.Text = Mid$(ddsDGAPPIS0(lIndex), 15, 30)
fgDetail.Col = 3
Select Case wField
    Case "DGAPPISSTA": fgDetail.Text = oldDGAPPIS0.DGAPPISSTA
    Case "DGAPPISVER": fgDetail.Text = oldDGAPPIS0.DGAPPISVER
    Case "DGAPPISPER": fgDetail.Text = oldDGAPPIS0.DGAPPISPER
    Case "DGAPPISETA": fgDetail.Text = oldDGAPPIS0.DGAPPISETA
    Case "DGAPPISSEQ": fgDetail.Text = oldDGAPPIS0.DGAPPISSEQ
    Case "DGAPPISKAM": fgDetail.Text = oldDGAPPIS0.DGAPPISKAM
    Case "DGAPPISCLI": fgDetail.Text = oldDGAPPIS0.DGAPPISCLI
    Case "DGAPPISDEC": fgDetail.Text = oldDGAPPIS0.DGAPPISDEC
    Case "DGAPPISMTE": fgDetail.Text = Format$(oldDGAPPIS0.DGAPPISMTE, "### ### ### ###.00")
    Case "DGAPPISNBJ": fgDetail.Text = oldDGAPPIS0.DGAPPISNBJ
    Case "DGAPPISVEC": fgDetail.Text = oldDGAPPIS0.DGAPPISVEC
    
    Case "GAPPISTAB": fgDetail.Text = oldDGAPPIS0.GAPPISTAB
    Case "GAPPISECH": fgDetail.Text = oldDGAPPIS0.GAPPISECH
    Case "GAPPISCLA": fgDetail.Text = oldDGAPPIS0.GAPPISCLA
    Case "GAPPISETA": fgDetail.Text = oldDGAPPIS0.GAPPISETA
    Case "GAPPISAGE": fgDetail.Text = oldDGAPPIS0.GAPPISAGE
    Case "GAPPISSER": fgDetail.Text = oldDGAPPIS0.GAPPISSER
    Case "GAPPISSSE": fgDetail.Text = oldDGAPPIS0.GAPPISSSE
    Case "GAPPISOPE": fgDetail.Text = oldDGAPPIS0.GAPPISOPE
    Case "GAPPISNAT": fgDetail.Text = oldDGAPPIS0.GAPPISNAT
    Case "GAPPISNUO": fgDetail.Text = oldDGAPPIS0.GAPPISNUO
    Case "GAPPISDEV": fgDetail.Text = oldDGAPPIS0.GAPPISDEV
    Case "GAPPISSEN": fgDetail.Text = oldDGAPPIS0.GAPPISSEN
    Case "GAPPISDEC": fgDetail.Text = oldDGAPPIS0.GAPPISDEC
    Case "GAPPISRUB": fgDetail.Text = oldDGAPPIS0.GAPPISRUB
    Case "GAPPISTPR": fgDetail.Text = oldDGAPPIS0.GAPPISTPR
    Case "GAPPISCLI": fgDetail.Text = oldDGAPPIS0.GAPPISCLI
    Case "GAPPISMON": fgDetail.Text = Format$(oldDGAPPIS0.GAPPISMON, "### ### ### ###.00")
    Case "GAPPISTTI": fgDetail.Text = oldDGAPPIS0.GAPPISTTI
    Case "GAPPISTTE": fgDetail.Text = oldDGAPPIS0.GAPPISTTE
    Case "GAPPISRTV": fgDetail.Text = oldDGAPPIS0.GAPPISRTV
    Case "GAPPISTAU": fgDetail.Text = Format$(oldDGAPPIS0.GAPPISTAU, "### ###.000 000")
    Case "GAPPISSOL": fgDetail.Text = Format$(oldDGAPPIS0.GAPPISSOL, "### ### ### ###.00")
    Case "GAPPISPOU": fgDetail.Text = Format$(oldDGAPPIS0.GAPPISPOU, "### ###..000 000")
    Case "GAPPISSIG": fgDetail.Text = oldDGAPPIS0.GAPPISSIG
    Case "GAPPISVIL": fgDetail.Text = oldDGAPPIS0.GAPPISVIL
    Case "GAPPISTMC": fgDetail.Text = Format$(oldDGAPPIS0.GAPPISTMC, "### ###..000 000")
    Case "GAPPISDAR": fgDetail.Text = oldDGAPPIS0.GAPPISDAR
    Case "GAPPISVAT": fgDetail.Text = Format$(oldDGAPPIS0.GAPPISVAT, "### ### ### ###.00")
    Case "GAPPISVAP": fgDetail.Text = Format$(oldDGAPPIS0.GAPPISVAP, "### ### ### ###.00")
    Case "GAPPISTVF": fgDetail.Text = oldDGAPPIS0.GAPPISTVF
    Case "GAPPISTP1": fgDetail.Text = Format$(oldDGAPPIS0.GAPPISTP1, "### ### ### ###.00")
    Case "GAPPISTP2": fgDetail.Text = Format$(oldDGAPPIS0.GAPPISTP2, "### ### ### ###.00")
    Case "GAPPISTM1": fgDetail.Text = Format$(oldDGAPPIS0.GAPPISTM1, "### ### ### ###.00")
    Case "GAPPISTM2": fgDetail.Text = Format$(oldDGAPPIS0.GAPPISTM2, "### ### ### ###.00")

End Select
If Mid$(ddsDGAPPIS0(lIndex), 2, 1) <> "*" Then fgDetail.CellBackColor = &HC0FFFF
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
    fgSelect.Visible = False
    fraDetail.Visible = False
    txtDetail.Visible = False
    cmdSelect_Ok.Visible = True
End If

End Sub


Private Sub cmdSelect_SQL()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdDGAPPIS0_SQL"
blnOk = False
   
Call DTPicker_Control(txtSelect_DGAPPISPER, wAmjMax)
xWhere = " where DGAPPISPER = " & wAmjMax

wCli = Val(txtSelect_DGAPPISCLI)
If wCli <> 0 Then blnOk = True: xWhere = xWhere & " and DGAPPISCLI = " & wCli

X = Trim(txtSelect_GAPPISTAB)
If X <> "" Then xWhere = xWhere & " and GAPPISTAB = " & X

X = Trim(txtSelect_GAPPISDEV)
If X <> "" Then blnOk = True: xWhere = xWhere & " and GAPPISDEV = '" & X & "'"

X = Trim(txtSelect_GAPPISSEN)
If X <> "" Then blnOk = True: xWhere = xWhere & " and GAPPISSEN = '" & X & "'"

X = Trim(txtSelect_GAPPISOPE)
If X <> "" Then blnOk = True: xWhere = xWhere & " and GAPPISOPE = '" & X & "'"

X = Trim(txtSelect_GAPPISNUO)
If X <> "" Then blnOk = True: xWhere = xWhere & " and GAPPISNUO = " & X

X = Trim(txtSelect_GAPPISRUB)
If X <> "" Then blnOk = True: xWhere = xWhere & " and GAPPISRUB = '" & X & "'"

If Not blnOk Then
    Call MsgBox("préciser plus de critères : n°client, code opération...", vbCritical, "BIA_DWH : DGAPPIS0")
    Exit Sub
End If

arrDGAPPIS0_SQL xWhere
fgSelect_Display

arrDGAPPISVEC(1) = "01M": arrDGAPPISVEC_AMJ(1) = dateElp("M-FM", 1, wAmjMax)
arrDGAPPISVEC(2) = "03M": arrDGAPPISVEC_AMJ(2) = dateElp("M-FM", 3, wAmjMax)
arrDGAPPISVEC(3) = "06M": arrDGAPPISVEC_AMJ(3) = dateElp("M-FM", 6, wAmjMax)
arrDGAPPISVEC(4) = "01A": arrDGAPPISVEC_AMJ(4) = dateElp("A-FM", 1, wAmjMax)
arrDGAPPISVEC(5) = "02A": arrDGAPPISVEC_AMJ(5) = dateElp("A-FM", 2, wAmjMax)
arrDGAPPISVEC(6) = "05A": arrDGAPPISVEC_AMJ(6) = dateElp("A-FM", 5, wAmjMax)
arrDGAPPISVEC(7) = "10A": arrDGAPPISVEC_AMJ(7) = dateElp("A-FM", 10, wAmjMax)
arrDGAPPISVEC(8) = "99A": arrDGAPPISVEC_AMJ(8) = 99999999


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub arrDGAPPIS0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrDGAPPIS0(101)
arrDGAPPIS0_Max = 100: arrDGAPPIS0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_BODWH & ".DGAPPIS0 " & xWhere & " order by DGAPPISCLI , GAPPISTAB , DGAPPISDEC"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsDGAPPIS0_GetBuffer(rsSab, xDGAPPIS0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDGAPPIS0.fgselect_Display"
        '' Exit Sub
     Else
         arrDGAPPIS0_Nb = arrDGAPPIS0_Nb + 1
         If arrDGAPPIS0_Nb > arrDGAPPIS0_Max Then
             arrDGAPPIS0_Max = arrDGAPPIS0_Max + 100
             ReDim Preserve arrDGAPPIS0(arrDGAPPIS0_Max)
         End If
         
         arrDGAPPIS0(arrDGAPPIS0_Nb) = xDGAPPIS0
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
cmdSelect_Ok.Visible = False
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgSelect_Display"
    
For I = 1 To arrDGAPPIS0_Nb
         
    xDGAPPIS0 = arrDGAPPIS0(I)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine I
Next I

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrDGAPPIS0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
On Error Resume Next


fgSelect.Col = 0: fgSelect.Text = xDGAPPIS0.DGAPPISSEQ
fgSelect.Col = 1: fgSelect.Text = xDGAPPIS0.DGAPPISKAM
fgSelect.Col = 2: fgSelect.Text = xDGAPPIS0.DGAPPISCLI
fgSelect.Col = 3: fgSelect.Text = xDGAPPIS0.GAPPISTAB
fgSelect.Col = 4: fgSelect.Text = dateImp10(xDGAPPIS0.DGAPPISDEC)
fgSelect.Col = 5: fgSelect.Text = xDGAPPIS0.GAPPISDEV

fgSelect.Col = 6: fgSelect.Text = Format$(xDGAPPIS0.GAPPISMON, "### ### ### ##0.00")
If xDGAPPIS0.GAPPISSEN = "A" Then
     fgSelect.CellForeColor = vbBlue
 Else
     fgSelect.CellForeColor = vbRed
 End If

fgSelect.Col = 7: fgSelect.Text = xDGAPPIS0.GAPPISSEN
fgSelect.Col = 8: fgSelect.Text = Format$(xDGAPPIS0.DGAPPISMTE, "### ### ### ##0.00")
If xDGAPPIS0.GAPPISSEN = "A" Then
     fgSelect.CellForeColor = vbBlue
 Else
     fgSelect.CellForeColor = vbRed
 End If
fgSelect.Col = 9: fgSelect.Text = xDGAPPIS0.GAPPISOPE & " " & xDGAPPIS0.GAPPISNAT & " " & xDGAPPIS0.GAPPISNUO


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
Call BiaPgmAut_Init(wFct, DGAPPIS0_Aut)

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
Call DTPicker_Set(txtSelect_DGAPPISPER, wAmjMax) '

ddsGAPPIS0_Init

txtSelect_GAPPISTAB.Clear
txtSelect_GAPPISTAB.AddItem " "
xSql = "select distinct GAPPISTAB from " & paramIBM_Library_BODWH & ".DGAPPIS0 order by GAPPISTAB"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    txtSelect_GAPPISTAB.AddItem Trim(rsSab("GAPPISTAB"))
    rsSab.MoveNext
Loop

txtSelect_GAPPISOPE.Clear
txtSelect_GAPPISOPE.AddItem " "
xSql = "select distinct GAPPISOPE from " & paramIBM_Library_BODWH & ".DGAPPIS0 order by GAPPISOPE"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    txtSelect_GAPPISOPE.AddItem Trim(rsSab("GAPPISOPE"))
    rsSab.MoveNext
Loop

txtSelect_GAPPISDEV.Clear
txtSelect_GAPPISDEV.AddItem " "
xSql = "select distinct GAPPISDEV from " & paramIBM_Library_BODWH & ".DGAPPIS0 order by GAPPISDEV"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    txtSelect_GAPPISDEV.AddItem Trim(rsSab("GAPPISDEV"))
    rsSab.MoveNext
Loop


txtSelect_GAPPISSEN.Clear
txtSelect_GAPPISSEN.AddItem " "
xSql = "select distinct GAPPISSEN from " & paramIBM_Library_BODWH & ".DGAPPIS0 order by GAPPISSEN"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    txtSelect_GAPPISSEN.AddItem Trim(rsSab("GAPPISSEN"))
    rsSab.MoveNext
Loop

Call MsgBox("- Les contrôles sur les champs modifiés sont très succints," & vbCrLf _
           & "- la cohérence des données n'est pas assûrée," & vbCrLf _
           & "- il n'y a pas de trace des modifications." & vbCrLf & vbCrLf _
           & "Les modifications du fichier DGAPPIS0 sont sous votre responsabilité !" _
           , vbExclamation, "maintenance du fichier BODWH / DGAPPIS0")
           
           

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
xDGAPPIS0 = newDGAPPIS0
xDGAPPIS0.DGAPPISSEQ = 900000000
Do
    xDGAPPIS0.DGAPPISSEQ = xDGAPPIS0.DGAPPISSEQ + 1
      
    If Not IsNull(sqlDGAPPIS0_Read(xDGAPPIS0)) Then blnOk = True
Loop Until blnOk
newDGAPPIS0.DGAPPISSEQ = xDGAPPIS0.DGAPPISSEQ
newDGAPPIS0.DGAPPISKAM = "+"

V = sqlDGAPPIS0_Insert(newDGAPPIS0)
fraDetail.Visible = False

cmdSelect_SQL
Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdDetail_Delete_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

V = sqlDGAPPIS0_Delete(oldDGAPPIS0)
fraDetail.Visible = False

cmdSelect_SQL
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdDetail_Quit_Click()
fraDetail.Visible = False
End Sub

Private Sub cmdDetail_Update_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
'If IsNull(cmdDGAPPIS0_Control) Then
    cmdDGAPPIS0_Update
    fraDetail.Visible = False
    cmdSelect_SQL
'End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdDGAPPIS0_Update()

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

V = sqlDGAPPIS0_Update(newDGAPPIS0, oldDGAPPIS0)

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
    
    cmdSelect_SQL
    
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
    If txtDetail_Update <> "*" Then
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
        fgSelect.Col = fgSelect_arrIndex:  arrDGAPPIS0_Index = CLng(fgSelect.Text)
        oldDGAPPIS0 = arrDGAPPIS0(arrDGAPPIS0_Index)
        newDGAPPIS0 = oldDGAPPIS0
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







Private Sub Text1_Change()

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


Private Sub txtSelect_DGAPPISCLI_Change()
cmdSelect_Reset

End Sub

Private Sub txtSelect_DGAPPISCLI_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub



Private Sub txtSelect_DGAPPISPER_Change()
cmdSelect_Reset

End Sub



Private Sub txtSelect_GAPPISDEV_Change()
cmdSelect_Reset

End Sub

Private Sub txtSelect_GAPPISDEV_Click()
cmdSelect_Reset

End Sub

Private Sub txtSelect_GAPPISDEV_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtSelect_GAPPISNUO_Change()
cmdSelect_Reset

End Sub

Private Sub txtSelect_GAPPISNUO_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub


Private Sub txtSelect_GAPPISOPE_Change()
cmdSelect_Reset

End Sub

Private Sub txtSelect_GAPPISOPE_Click()
cmdSelect_Reset

End Sub

Private Sub txtSelect_GAPPISOPE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtSelect_GAPPISRUB_Change()
cmdSelect_Reset

End Sub

Private Sub txtSelect_GAPPISRUB_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub


Private Sub txtSelect_GAPPISSEN_Change()
cmdSelect_Reset
End Sub

Private Sub txtSelect_GAPPISSEN_Click()
cmdSelect_Reset

End Sub

Private Sub txtSelect_GAPPISSEN_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSelect_GAPPISTAB_Change()
cmdSelect_Reset

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
        Case "DGAPPISCLI": xDGAPPIS0.DGAPPISCLI = txtDetail_N
                         If oldDGAPPIS0.DGAPPISCLI = xDGAPPIS0.DGAPPISCLI Then
                            newDGAPPIS0.DGAPPISCLI = oldDGAPPIS0.DGAPPISCLI
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            X = fraDetail_Control_DGAPPISCLI(xDGAPPIS0.DGAPPISCLI)
                            If X = "" Then
                                Call MsgBox("client inconnu : " & xDGAPPIS0.DGAPPISCLI, vbExclamation, "DGAPPIS0 contrôle client")
                            Else
                                Call lstErr_Clear(lstErr, cmdContext, xDGAPPIS0.DGAPPISCLI & " : " & X): DoEvents
                                newDGAPPIS0.DGAPPISCLI = xDGAPPIS0.DGAPPISCLI
                                fgDetail = newDGAPPIS0.DGAPPISCLI: fgDetail.CellBackColor = txtDetail_ColorUpdate
                            End If
                        End If
         Case "DGAPPISDEC": xDGAPPIS0.DGAPPISDEC = txtDetail_N
                         If oldDGAPPIS0.DGAPPISDEC = xDGAPPIS0.DGAPPISDEC Then
                            newDGAPPIS0.DGAPPISDEC = oldDGAPPIS0.DGAPPISDEC
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                            
                            fgDetail.Row = fgDetail.Row + 2
                            newDGAPPIS0.DGAPPISNBJ = oldDGAPPIS0.DGAPPISNBJ
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                            fgDetail.Row = fgDetail.Row + 1
                            newDGAPPIS0.DGAPPISVEC = oldDGAPPIS0.DGAPPISVEC
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                       Else
                            X = dateElp("DateMS", 0, xDGAPPIS0.DGAPPISDEC)
                            If X = "" Then
                                Call MsgBox("date non conforme 'aaaammjj' : " & xDGAPPIS0.DGAPPISDEC, vbExclamation, "DGAPPIS0 contrôle échéance")
                            Else
                                newDGAPPIS0.DGAPPISDEC = xDGAPPIS0.DGAPPISDEC
                                fgDetail = newDGAPPIS0.DGAPPISDEC: fgDetail.CellBackColor = txtDetail_ColorUpdate
                                newDGAPPIS0.DGAPPISNBJ = oldDGAPPIS0.DGAPPISNBJ
                                
                                fgDetail.Row = fgDetail.Row + 2
                                newDGAPPIS0.DGAPPISNBJ = DateDiff("d", Format(newDGAPPIS0.DGAPPISPER, "@@@@/@@/@@"), X)
                                fgDetail = newDGAPPIS0.DGAPPISNBJ: fgDetail.CellBackColor = txtDetail_ColorUpdate
                                fgDetail.Row = fgDetail.Row + 1
                                newDGAPPIS0.DGAPPISVEC = fraDetail_Control_DGAPPISVEC(newDGAPPIS0.DGAPPISDEC)
                                fgDetail = newDGAPPIS0.DGAPPISVEC: fgDetail.CellBackColor = txtDetail_ColorUpdate
                            End If
                       End If
       Case "DGAPPISMTE": xDGAPPIS0.DGAPPISMTE = txtDetail_C
                         If oldDGAPPIS0.DGAPPISMTE = xDGAPPIS0.DGAPPISMTE Then
                            newDGAPPIS0.DGAPPISMTE = oldDGAPPIS0.DGAPPISMTE
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.DGAPPISMTE = xDGAPPIS0.DGAPPISMTE
                            fgDetail = Format$(newDGAPPIS0.DGAPPISMTE, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISTAB": xDGAPPIS0.GAPPISTAB = txtDetail_N
                        If oldDGAPPIS0.GAPPISTAB = xDGAPPIS0.GAPPISTAB Then
                            newDGAPPIS0.GAPPISTAB = oldDGAPPIS0.GAPPISTAB
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISTAB = xDGAPPIS0.GAPPISTAB
                            fgDetail = newDGAPPIS0.GAPPISTAB: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISECH": xDGAPPIS0.GAPPISECH = txtDetail_N
                        If oldDGAPPIS0.GAPPISECH = xDGAPPIS0.GAPPISECH Then
                            newDGAPPIS0.GAPPISECH = oldDGAPPIS0.GAPPISECH
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISECH = xDGAPPIS0.GAPPISECH
                            fgDetail = newDGAPPIS0.GAPPISECH: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISCLA": xDGAPPIS0.GAPPISCLA = txtDetail_N
                        If oldDGAPPIS0.GAPPISCLA = xDGAPPIS0.GAPPISCLA Then
                            newDGAPPIS0.GAPPISCLA = oldDGAPPIS0.GAPPISCLA
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISCLA = xDGAPPIS0.GAPPISCLA
                            fgDetail = newDGAPPIS0.GAPPISCLA: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISDEV": xDGAPPIS0.GAPPISDEV = txtDetail_A
                        If oldDGAPPIS0.GAPPISDEV = xDGAPPIS0.GAPPISDEV Then
                            newDGAPPIS0.GAPPISDEV = oldDGAPPIS0.GAPPISDEV
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                             X = fraDetail_Control_GAPPISDEV(xDGAPPIS0.GAPPISDEV)
                            If X = "" Then
                                Call MsgBox("devise inconnue : " & xDGAPPIS0.GAPPISDEV, vbExclamation, "DGAPPIS0 contrôle devise")
                            Else
                                Call lstErr_Clear(lstErr, cmdContext, xDGAPPIS0.GAPPISDEV & " : " & X): DoEvents
                                newDGAPPIS0.GAPPISDEV = xDGAPPIS0.GAPPISDEV
                                fgDetail = newDGAPPIS0.GAPPISDEV: fgDetail.CellBackColor = txtDetail_ColorUpdate
                            End If
                        End If
       Case "GAPPISSEN": xDGAPPIS0.GAPPISSEN = txtDetail_A
                        If oldDGAPPIS0.GAPPISSEN = xDGAPPIS0.GAPPISSEN Then
                            newDGAPPIS0.GAPPISSEN = oldDGAPPIS0.GAPPISSEN
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISSEN = xDGAPPIS0.GAPPISSEN
                            fgDetail = newDGAPPIS0.GAPPISSEN: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISRUB": xDGAPPIS0.GAPPISRUB = txtDetail_A
                        If oldDGAPPIS0.GAPPISRUB = xDGAPPIS0.GAPPISRUB Then
                           newDGAPPIS0.GAPPISRUB = oldDGAPPIS0.GAPPISRUB
                           fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISRUB = xDGAPPIS0.GAPPISRUB
                            fgDetail = newDGAPPIS0.GAPPISRUB: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISTPR": xDGAPPIS0.GAPPISTPR = txtDetail_A
                        If oldDGAPPIS0.GAPPISTPR = xDGAPPIS0.GAPPISTPR Then
                            newDGAPPIS0.GAPPISTPR = oldDGAPPIS0.GAPPISTPR
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISTPR = xDGAPPIS0.GAPPISTPR
                            fgDetail = newDGAPPIS0.GAPPISTPR: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISMON": xDGAPPIS0.GAPPISMON = txtDetail_C
                        If oldDGAPPIS0.GAPPISMON = xDGAPPIS0.GAPPISMON Then
                            newDGAPPIS0.GAPPISMON = oldDGAPPIS0.GAPPISMON
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISMON = xDGAPPIS0.GAPPISMON
                            fgDetail = Format$(newDGAPPIS0.GAPPISMON, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISTTI": xDGAPPIS0.GAPPISTTI = txtDetail_A
                        If oldDGAPPIS0.GAPPISTTI = xDGAPPIS0.GAPPISTTI Then
                            newDGAPPIS0.GAPPISTTI = oldDGAPPIS0.GAPPISTTI
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISTTI = xDGAPPIS0.GAPPISTTI
                            fgDetail = newDGAPPIS0.GAPPISTTI: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISTTE": xDGAPPIS0.GAPPISTTE = txtDetail_A
                        If oldDGAPPIS0.GAPPISTTE = xDGAPPIS0.GAPPISTTE Then
                            newDGAPPIS0.GAPPISTTE = oldDGAPPIS0.GAPPISTTE
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISTTE = xDGAPPIS0.GAPPISTTE
                            fgDetail = newDGAPPIS0.GAPPISTTE: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISRTV": xDGAPPIS0.GAPPISRTV = txtDetail_A
                        If oldDGAPPIS0.GAPPISRTV = xDGAPPIS0.GAPPISRTV Then
                            newDGAPPIS0.GAPPISRTV = oldDGAPPIS0.GAPPISRTV
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISRTV = xDGAPPIS0.GAPPISRTV
                            fgDetail = newDGAPPIS0.GAPPISRTV: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISTAU": xDGAPPIS0.GAPPISTAU = txtDetail_D
                        If oldDGAPPIS0.GAPPISTAU = xDGAPPIS0.GAPPISTAU Then
                            newDGAPPIS0.GAPPISTAU = oldDGAPPIS0.GAPPISTAU
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISTAU = xDGAPPIS0.GAPPISTAU
                            fgDetail = Format$(newDGAPPIS0.GAPPISTAU, "### ###.000 000"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISSOL": xDGAPPIS0.GAPPISSOL = txtDetail_C
                        If oldDGAPPIS0.GAPPISSOL = xDGAPPIS0.GAPPISSOL Then
                            newDGAPPIS0.GAPPISSOL = oldDGAPPIS0.GAPPISSOL
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISSOL = xDGAPPIS0.GAPPISSOL
                            fgDetail = Format$(newDGAPPIS0.GAPPISSOL, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISPOU": xDGAPPIS0.GAPPISPOU = txtDetail_D
                        If oldDGAPPIS0.GAPPISPOU = xDGAPPIS0.GAPPISPOU Then
                            newDGAPPIS0.GAPPISPOU = oldDGAPPIS0.GAPPISPOU
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISPOU = xDGAPPIS0.GAPPISPOU
                            fgDetail = Format$(newDGAPPIS0.GAPPISPOU, "### ###.000 000"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISSIG": xDGAPPIS0.GAPPISSIG = txtDetail_A
                        If oldDGAPPIS0.GAPPISSIG = xDGAPPIS0.GAPPISSIG Then
                            newDGAPPIS0.GAPPISSIG = oldDGAPPIS0.GAPPISSIG
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISSIG = xDGAPPIS0.GAPPISSIG
                            fgDetail = newDGAPPIS0.GAPPISSIG: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISVIL": xDGAPPIS0.GAPPISVIL = txtDetail_A
                        If oldDGAPPIS0.GAPPISVIL = xDGAPPIS0.GAPPISVIL Then
                            newDGAPPIS0.GAPPISVIL = oldDGAPPIS0.GAPPISVIL
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISVIL = xDGAPPIS0.GAPPISVIL
                            fgDetail = newDGAPPIS0.GAPPISVIL: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISTMC": xDGAPPIS0.GAPPISTMC = txtDetail_D
                        If oldDGAPPIS0.GAPPISTMC = xDGAPPIS0.GAPPISTMC Then
                            newDGAPPIS0.GAPPISTMC = oldDGAPPIS0.GAPPISTMC
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISTMC = xDGAPPIS0.GAPPISTMC
                            fgDetail = Format$(newDGAPPIS0.GAPPISTMC, "### ###.000 000"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISDAR": xDGAPPIS0.GAPPISDAR = txtDetail_N
                        If oldDGAPPIS0.GAPPISDAR = xDGAPPIS0.GAPPISDAR Then
                            newDGAPPIS0.GAPPISDAR = oldDGAPPIS0.GAPPISDAR
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISDAR = xDGAPPIS0.GAPPISDAR
                            fgDetail = newDGAPPIS0.GAPPISDAR: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISVAT": xDGAPPIS0.GAPPISVAT = txtDetail_C
                        If oldDGAPPIS0.GAPPISVAT = xDGAPPIS0.GAPPISVAT Then
                            newDGAPPIS0.GAPPISVAT = oldDGAPPIS0.GAPPISVAT
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISVAT = xDGAPPIS0.GAPPISVAT
                            fgDetail = Format$(newDGAPPIS0.GAPPISVAT, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISVAP": xDGAPPIS0.GAPPISVAP = txtDetail_C
                        If oldDGAPPIS0.GAPPISVAP = xDGAPPIS0.GAPPISVAP Then
                            newDGAPPIS0.GAPPISVAP = oldDGAPPIS0.GAPPISVAP
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISVAP = xDGAPPIS0.GAPPISVAP
                            fgDetail = Format$(newDGAPPIS0.GAPPISVAP, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISTVF": xDGAPPIS0.GAPPISTVF = txtDetail_A
                        If oldDGAPPIS0.GAPPISTVF = xDGAPPIS0.GAPPISTVF Then
                            newDGAPPIS0.GAPPISTVF = oldDGAPPIS0.GAPPISTVF
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISTVF = xDGAPPIS0.GAPPISTVF
                            fgDetail = newDGAPPIS0.GAPPISTVF: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISTP1": xDGAPPIS0.GAPPISTP1 = txtDetail_C
                        If oldDGAPPIS0.GAPPISTP1 = xDGAPPIS0.GAPPISTP1 Then
                            newDGAPPIS0.GAPPISTP1 = oldDGAPPIS0.GAPPISTP1
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISTP1 = xDGAPPIS0.GAPPISTP1
                            fgDetail = Format$(newDGAPPIS0.GAPPISTP1, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISTP2": xDGAPPIS0.GAPPISTP2 = txtDetail_C
                        If oldDGAPPIS0.GAPPISTP2 = xDGAPPIS0.GAPPISTP2 Then
                            newDGAPPIS0.GAPPISTP2 = oldDGAPPIS0.GAPPISTP2
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISTP2 = xDGAPPIS0.GAPPISTP2
                            fgDetail = Format$(newDGAPPIS0.GAPPISTP2, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISTM1": xDGAPPIS0.GAPPISTM1 = txtDetail_C
                        If oldDGAPPIS0.GAPPISTM1 = xDGAPPIS0.GAPPISTM1 Then
                            newDGAPPIS0.GAPPISTM1 = oldDGAPPIS0.GAPPISTM1
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISTM1 = xDGAPPIS0.GAPPISTM1
                            fgDetail = Format$(newDGAPPIS0.GAPPISTM1, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "GAPPISTM2": xDGAPPIS0.GAPPISTM2 = txtDetail_C
                        If oldDGAPPIS0.GAPPISTM2 = xDGAPPIS0.GAPPISTM2 Then
                            newDGAPPIS0.GAPPISTM2 = oldDGAPPIS0.GAPPISTM2
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDGAPPIS0.GAPPISTM2 = xDGAPPIS0.GAPPISTM2
                            fgDetail = Format$(newDGAPPIS0.GAPPISTM2, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
End Select
fgDetail.Row = fgDetail_RowClick
txtDetail_blnUpdate = False: cmdDetail_Control.Visible = False
cmdDetail_Update.Visible = DGAPPIS0_Aut.Saisir
End Sub

Public Function fraDetail_Control_DGAPPISVEC(lAMJ As Long)
Dim K As Integer
For K = 1 To 8
    If lAMJ <= arrDGAPPISVEC_AMJ(K) Then Exit For
Next K
fraDetail_Control_DGAPPISVEC = arrDGAPPISVEC(K)
End Function
Public Function fraDetail_Control_DGAPPISCLI(lDGAPPISCLI As Long)
Dim xSql As String
fraDetail_Control_DGAPPISCLI = ""
xSql = "select CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & Format$(lDGAPPISCLI, "0000000") & "'"
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then fraDetail_Control_DGAPPISCLI = rsSab("CLIENARA1")

End Function

Public Function fraDetail_Control_GAPPISDEV(lGAPPISDEV As String)
Dim xSql As String
fraDetail_Control_GAPPISDEV = ""
xSql = "select BASDVSABR from " & paramIBM_Library_SAB & ".ZBASDVS0 where BASDVSDEV = '" & lGAPPISDEV & "'"
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then fraDetail_Control_GAPPISDEV = rsSab("BASDVSABR")

End Function

Private Sub txtSelect_GAPPISTAB_Click()
cmdSelect_Reset

End Sub

Private Sub txtSelect_GAPPISTAB_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub

