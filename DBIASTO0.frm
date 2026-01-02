VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDBIASTO0 
   AutoRedraw      =   -1  'True
   Caption         =   "DBIASTO0 : maintenance"
   ClientHeight    =   9312
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   13572
   Icon            =   "DBIASTO0.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9312
   ScaleWidth      =   13572
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   240
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
      _ExtentY        =   15325
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "DGAPPIS0"
      TabPicture(0)   =   "DBIASTO0.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "......."
      TabPicture(1)   =   "DBIASTO0.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraDetail"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraDetail 
         BackColor       =   &H00A0E0FF&
         Height          =   6852
         Left            =   2640
         TabIndex        =   23
         Top             =   600
         Width           =   7500
         Begin VB.CommandButton cmdDetail_Control 
            BackColor       =   &H0080C0FF&
            Caption         =   "Saisie"
            Height          =   500
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   6120
            Width           =   1095
         End
         Begin VB.CommandButton cmdDetail_Quit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Abandonner"
            Enabled         =   0   'False
            Height          =   500
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   6150
            Width           =   1095
         End
         Begin VB.CommandButton cmdDetail_Copy 
            BackColor       =   &H00FF80FF&
            Caption         =   "Copier"
            Height          =   500
            Left            =   1680
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
         Begin VB.CommandButton cmdDetail_Update 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Enregistrer"
            Height          =   500
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   6150
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   5865
            Left            =   135
            TabIndex        =   29
            Top             =   100
            Width           =   7200
            _ExtentX        =   12700
            _ExtentY        =   10351
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
               Size            =   7.8
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
         Left            =   -74880
         TabIndex        =   3
         Top             =   600
         Width           =   13296
         Begin VB.TextBox txtDetail 
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.8
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
            Begin VB.ComboBox txtSelect_YSTOPCI 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   324
               Left            =   9240
               Sorted          =   -1  'True
               TabIndex        =   22
               Text            =   "txtSelect_YSTOPCI"
               Top             =   480
               Width           =   1335
            End
            Begin VB.ComboBox txtSelect_YSTONAT 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   324
               Left            =   7560
               Sorted          =   -1  'True
               TabIndex        =   20
               Text            =   "txtSelect_YSTONAT"
               Top             =   480
               Width           =   1335
            End
            Begin VB.ComboBox txtSelect_YSTODEV 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   324
               Left            =   3360
               Sorted          =   -1  'True
               TabIndex        =   16
               Text            =   "txtSelect_GAPPISDEV"
               Top             =   500
               Width           =   852
            End
            Begin VB.ComboBox txtSelect_YSTOOPE 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   324
               Left            =   5880
               Sorted          =   -1  'True
               TabIndex        =   15
               Text            =   "txtSelect_YSTOOPE"
               Top             =   480
               Width           =   1332
            End
            Begin VB.ComboBox txtSelect_YSTOAPP 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   324
               Left            =   4560
               Sorted          =   -1  'True
               TabIndex        =   14
               Text            =   "txtSelect_YSTOAPP"
               Top             =   480
               Width           =   852
            End
            Begin VB.TextBox txtSelect_DBIASTOCLI 
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   324
               Left            =   1680
               TabIndex        =   11
               Top             =   500
               Width           =   1212
            End
            Begin MSComCtl2.DTPicker txtSelect_DBIASTOPER 
               Height          =   300
               Left            =   120
               TabIndex        =   7
               Top             =   504
               Width           =   1332
               _ExtentX        =   2350
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.8
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
            Begin VB.Label lblSelect_YSTOPCI 
               BackColor       =   &H00F0FFFF&
               Caption         =   "PCI"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   9600
               TabIndex        =   21
               Top             =   240
               Width           =   612
            End
            Begin VB.Label lblSelect_YSTONAT 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Nature"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   7920
               TabIndex        =   19
               Top             =   240
               Width           =   612
            End
            Begin VB.Label lblSelect_YSTODEV 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Devise"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   3480
               TabIndex        =   18
               Top             =   240
               Width           =   612
            End
            Begin VB.Label lblSelect_YSTOOPE 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Opération"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   6120
               TabIndex        =   17
               Top             =   240
               Width           =   852
            End
            Begin VB.Label lblSelect_YSTOAPP 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Application"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   4560
               TabIndex        =   12
               Top             =   240
               Width           =   852
            End
            Begin VB.Label lblSelect_DBIASTOCLI 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Client"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   2040
               TabIndex        =   10
               Top             =   240
               Width           =   612
            End
            Begin VB.Label lblSelect_DBIASTOPER 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Période"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   480
               TabIndex        =   9
               Top             =   240
               Width           =   612
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6825
            Left            =   240
            TabIndex        =   5
            Top             =   1320
            Width           =   12555
            _ExtentX        =   22140
            _ExtentY        =   12044
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
            FormatString    =   $"DBIASTO0.frx":0342
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.4
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
      Picture         =   "DBIASTO0.frx":041C
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
Attribute VB_Name = "frmDBIASTO0"
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
Dim DBIASTO0_Aut As typeAuthorization
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
Dim xDBIASTO0 As typeDBIASTO0, newDBIASTO0 As typeDBIASTO0, oldDBIASTO0 As typeDBIASTO0
Dim arrDBIASTO0() As typeDBIASTO0, arrDBIASTO0_Nb As Long, arrDBIASTO0_Max As Long, arrDBIASTO0_Index As Long

Dim txtDetail_Type As String, txtDetail_Update As String
Dim txtDetail_Field As String
Dim txtDetail_blnUpdate As Boolean, txtDetail_ColorUpdate As Long
Dim txtDetail_Row As Integer

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
cmdDetail_Delete.Visible = DBIASTO0_Aut.Saisir
cmdDetail_Copy.Visible = DBIASTO0_Aut.Saisir
cmdDetail_Update.Visible = False
cmdDetail_Quit.Enabled = True
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString

    
For I = 1 To 33
         
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



wField = Trim(Mid$(ddsDBIASTO0(lIndex), 4, 10))
fgDetail.Col = 0: fgDetail.Text = Mid$(ddsDBIASTO0(lIndex), 1, 2)
fgDetail.Col = 1: fgDetail.Text = wField
fgDetail.Col = 2: fgDetail.Text = Mid$(ddsDBIASTO0(lIndex), 15, 30)
fgDetail.Col = 3
Select Case wField
    Case "DBIASTOSTA": fgDetail.Text = oldDBIASTO0.DBIASTOSTA
    Case "DBIASTOVER": fgDetail.Text = oldDBIASTO0.DBIASTOVER
    Case "DBIASTOPER": fgDetail.Text = oldDBIASTO0.DBIASTOPER
    Case "DBIASTOETA": fgDetail.Text = oldDBIASTO0.DBIASTOETA
    Case "DBIASTOSEQ": fgDetail.Text = oldDBIASTO0.DBIASTOSEQ
    Case "DBIASTOKAM": fgDetail.Text = oldDBIASTO0.DBIASTOKAM
    Case "DBIASTOCLI": fgDetail.Text = oldDBIASTO0.DBIASTOCLI
    Case "DBIASTOAUT": fgDetail.Text = oldDBIASTO0.DBIASTOAUT
    Case "DBIASTOMTE": fgDetail.Text = Format$(oldDBIASTO0.DBIASTOMTE, "### ### ### ###.00")
    Case "DBIASTOAU0": fgDetail.Text = oldDBIASTO0.DBIASTOAU0
    Case "DBIASTOCPT": fgDetail.Text = oldDBIASTO0.DBIASTOCPT
    
    Case "YSTOETA": fgDetail.Text = oldDBIASTO0.YSTOETA
    Case "YSTOAGE": fgDetail.Text = oldDBIASTO0.YSTOAGE
    Case "YSTOSER": fgDetail.Text = oldDBIASTO0.YSTOSER
    Case "YSTOSSE": fgDetail.Text = oldDBIASTO0.YSTOSSE
    Case "YSTOOPE": fgDetail.Text = oldDBIASTO0.YSTOOPE
    Case "YSTONUM": fgDetail.Text = oldDBIASTO0.YSTONUM
    Case "YSTOSEQ": fgDetail.Text = oldDBIASTO0.YSTOSEQ
    Case "YSTOPCI": fgDetail.Text = oldDBIASTO0.YSTOPCI
    Case "YSTOCCL": fgDetail.Text = oldDBIASTO0.YSTOCCL
    Case "YSTOCLI": fgDetail.Text = oldDBIASTO0.YSTOCLI
    Case "YSTODEV": fgDetail.Text = oldDBIASTO0.YSTODEV
    Case "YSTOMON": fgDetail.Text = Format$(oldDBIASTO0.YSTOMON, "### ### ### ###.00")
    Case "YSTODEB": fgDetail.Text = oldDBIASTO0.YSTODEB
    Case "YSTOFIN": fgDetail.Text = oldDBIASTO0.YSTOFIN
    Case "YSTOAPP": fgDetail.Text = oldDBIASTO0.YSTOAPP
    Case "YSTONAT": fgDetail.Text = oldDBIASTO0.YSTONAT
    Case "YSTOCC1": fgDetail.Text = oldDBIASTO0.YSTOCC1
    Case "YSTOCL1": fgDetail.Text = oldDBIASTO0.YSTOCL1
    Case "YSTOCC2": fgDetail.Text = oldDBIASTO0.YSTOCC2
    Case "YSTOCL2": fgDetail.Text = oldDBIASTO0.YSTOCL2
    Case "YSTOCTX": fgDetail.Text = oldDBIASTO0.YSTOCTX
    Case "YSTOTAU": fgDetail.Text = Format$(oldDBIASTO0.YSTOTAU, "### ###.000 000")


End Select
If Mid$(ddsDBIASTO0(lIndex), 2, 1) <> "*" Then fgDetail.CellBackColor = &HC0FFFF
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

currentAction = "cmdDBIASTO0_SQL"
blnOk = False
   
Call DTPicker_Control(txtSelect_DBIASTOPER, wAmjMax)
xWhere = " where DBIASTOPER = " & wAmjMax

wCli = Val(txtSelect_DBIASTOCLI)
If wCli <> 0 Then blnOk = True: xWhere = xWhere & " and DBIASTOCLI = " & wCli

X = Trim(txtSelect_YSTOAPP)
If X <> "" Then blnOk = True: xWhere = xWhere & " and YSTOAPP = '" & X & "'"

X = Trim(txtSelect_YSTODEV)
If X <> "" Then blnOk = True: xWhere = xWhere & " and YSTODEV = '" & X & "'"

X = Trim(txtSelect_YSTONAT)
If X <> "" Then blnOk = True: xWhere = xWhere & " and YSTONAT = '" & X & "'"

X = Trim(txtSelect_YSTOOPE)
If X <> "" Then blnOk = True: xWhere = xWhere & " and YSTOOPE = '" & X & "'"

X = Trim(txtSelect_YSTOPCI)
If X <> "" Then blnOk = True: xWhere = xWhere & " and YSTOPCI like '" & X & "%'"

If Not blnOk Then
    Call MsgBox("préciser plus de critères : n°client, code opération...", vbCritical, "BIA_DWH : DBIASTO0")
    Exit Sub
End If

arrDBIASTO0_SQL xWhere
fgSelect_Display


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub arrDBIASTO0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrDBIASTO0(101)
arrDBIASTO0_Max = 100: arrDBIASTO0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_BODWH & ".DBIASTO0 " & xWhere & " order by DBIASTOCLI , YSTOAPP , DBIASTOSEQ"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsDBIASTO0_GetBuffer(rsSab, xDBIASTO0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDBIASTO0.fgselect_Display"
        '' Exit Sub
     Else
         arrDBIASTO0_Nb = arrDBIASTO0_Nb + 1
         If arrDBIASTO0_Nb > arrDBIASTO0_Max Then
             arrDBIASTO0_Max = arrDBIASTO0_Max + 100
             ReDim Preserve arrDBIASTO0(arrDBIASTO0_Max)
         End If
         
         arrDBIASTO0(arrDBIASTO0_Nb) = xDBIASTO0
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
    
For I = 1 To arrDBIASTO0_Nb
         
    xDBIASTO0 = arrDBIASTO0(I)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine I
Next I

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrDBIASTO0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
On Error Resume Next


fgSelect.Col = 0: fgSelect.Text = xDBIASTO0.DBIASTOSEQ
fgSelect.Col = 1: fgSelect.Text = xDBIASTO0.DBIASTOKAM
fgSelect.Col = 2: fgSelect.Text = xDBIASTO0.DBIASTOCLI
fgSelect.Col = 3: fgSelect.Text = xDBIASTO0.YSTOAPP
fgSelect.Col = 4: fgSelect.Text = Trim(xDBIASTO0.DBIASTOCPT)
fgSelect.Col = 5: fgSelect.Text = xDBIASTO0.YSTODEV

fgSelect.Col = 6: fgSelect.Text = Format$(xDBIASTO0.YSTOMON, "### ### ### ##0.00")

fgSelect.Col = 7: fgSelect.Text = Format$(xDBIASTO0.DBIASTOMTE, "### ### ### ##0.00")
fgSelect.Col = 8: fgSelect.Text = xDBIASTO0.YSTOOPE & " " & xDBIASTO0.YSTONAT & " " & xDBIASTO0.YSTONUM & " " & xDBIASTO0.YSTOSEQ
fgSelect.Col = 9: fgSelect.Text = Trim(xDBIASTO0.DBIASTOAUT)
fgSelect.Col = 10: fgSelect.Text = Trim(xDBIASTO0.DBIASTOAU0)


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
Call BiaPgmAut_Init(wFct, DBIASTO0_Aut)

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
Call DTPicker_Set(txtSelect_DBIASTOPER, wAmjMax) '

ddsYSTO0_Init

txtSelect_YSTOAPP.Clear
txtSelect_YSTOAPP.AddItem " "
xSql = "select distinct YSTOAPP from " & paramIBM_Library_BODWH & ".DBIASTO0 order by YSTOAPP"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    txtSelect_YSTOAPP.AddItem Trim(rsSab("YSTOAPP"))
    rsSab.MoveNext
Loop

txtSelect_YSTOOPE.Clear
txtSelect_YSTOOPE.AddItem " "
xSql = "select distinct YSTOOPE from " & paramIBM_Library_BODWH & ".DBIASTO0 order by YSTOOPE"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    txtSelect_YSTOOPE.AddItem Trim(rsSab("YSTOOPE"))
    rsSab.MoveNext
Loop

txtSelect_YSTODEV.Clear
txtSelect_YSTODEV.AddItem " "
xSql = "select distinct YSTODEV from " & paramIBM_Library_BODWH & ".DBIASTO0 order by YSTODEV"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    txtSelect_YSTODEV.AddItem Trim(rsSab("YSTODEV"))
    rsSab.MoveNext
Loop


txtSelect_YSTONAT.Clear
txtSelect_YSTONAT.AddItem " "
xSql = "select distinct YSTONAT from " & paramIBM_Library_BODWH & ".DBIASTO0 order by YSTONAT"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    txtSelect_YSTONAT.AddItem Trim(rsSab("YSTONAT"))
    rsSab.MoveNext
Loop

txtSelect_YSTOPCI.Clear
txtSelect_YSTOPCI.AddItem " "
xSql = "select distinct YSTOPCI from " & paramIBM_Library_BODWH & ".DBIASTO0 order by YSTOPCI"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    txtSelect_YSTOPCI.AddItem Trim(rsSab("YSTOPCI"))
    rsSab.MoveNext
Loop

Call MsgBox("- Les contrôles sur les champs modifiés sont très succints," & vbCrLf _
           & "- la cohérence des données n'est pas assûrée," & vbCrLf _
           & "- il n'y a pas de trace des modifications." & vbCrLf & vbCrLf _
           & "Les modifications du fichier DBIASTO0 sont sous votre responsabilité !" _
           , vbExclamation, "maintenance du fichier BODWH / DBIASTO0")
           
           

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
xDBIASTO0 = newDBIASTO0
xDBIASTO0.DBIASTOSEQ = 900000000
Do
    xDBIASTO0.DBIASTOSEQ = xDBIASTO0.DBIASTOSEQ + 1
      
    If Not IsNull(sqlDBIASTO0_Read(xDBIASTO0)) Then blnOk = True
Loop Until blnOk
newDBIASTO0.DBIASTOSEQ = xDBIASTO0.DBIASTOSEQ
newDBIASTO0.DBIASTOKAM = "+"

V = sqlDBIASTO0_Insert(newDBIASTO0)
fraDetail.Visible = False

cmdSelect_SQL
Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdDetail_Delete_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

V = sqlDBIASTO0_Delete(oldDBIASTO0)
fraDetail.Visible = False

cmdSelect_SQL
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdDetail_Quit_Click()
fraDetail.Visible = False
End Sub

Private Sub cmdDetail_Update_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
'If IsNull(cmdDBIASTO0_Control) Then
    cmdDBIASTO0_Update
    fraDetail.Visible = False
    cmdSelect_SQL
'End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdDBIASTO0_Update()

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

V = sqlDBIASTO0_Update(newDBIASTO0, oldDBIASTO0)

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
        fgSelect.Col = fgSelect_arrIndex:  arrDBIASTO0_Index = CLng(fgSelect.Text)
        oldDBIASTO0 = arrDBIASTO0(arrDBIASTO0_Index)
        newDBIASTO0 = oldDBIASTO0
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


Private Sub txtSelect_DBIASTOCLI_Change()
cmdSelect_Reset

End Sub

Private Sub txtSelect_DBIASTOCLI_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtSelect_DBIASTOPER_Change()
cmdSelect_Reset

End Sub


Private Sub txtSelect_YSTODEV_Change()
cmdSelect_Reset

End Sub

Private Sub txtSelect_YSTODEV_Click()
cmdSelect_Reset

End Sub

Private Sub txtSelect_YSTODEV_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtSelect_YSTOOPE_Change()
cmdSelect_Reset

End Sub

Private Sub txtSelect_YSTOOPE_Click()
cmdSelect_Reset

End Sub

Private Sub txtSelect_YSTOOPE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtSelect_YSTONAT_Change()
cmdSelect_Reset
End Sub

Private Sub txtSelect_YSTONAT_Click()
cmdSelect_Reset

End Sub

Private Sub txtSelect_YSTONAT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSelect_YSTOAPP_Change()
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
        Case "DBIASTOCLI": xDBIASTO0.DBIASTOCLI = txtDetail_N
                         If oldDBIASTO0.DBIASTOCLI = xDBIASTO0.DBIASTOCLI Then
                            newDBIASTO0.DBIASTOCLI = oldDBIASTO0.DBIASTOCLI
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            X = fraDetail_Control_DBIASTOCLI(xDBIASTO0.DBIASTOCLI)
                            If X = "" Then
                                Call MsgBox("client inconnu : " & xDBIASTO0.DBIASTOCLI, vbExclamation, "DBIASTO0 contrôle client")
                            Else
                                Call lstErr_Clear(lstErr, cmdContext, xDBIASTO0.DBIASTOCLI & " : " & X): DoEvents
                                newDBIASTO0.DBIASTOCLI = xDBIASTO0.DBIASTOCLI
                                fgDetail = newDBIASTO0.DBIASTOCLI: fgDetail.CellBackColor = txtDetail_ColorUpdate
                            End If
                        End If
       Case "DBIASTOMTE": xDBIASTO0.DBIASTOMTE = txtDetail_C
                         If oldDBIASTO0.DBIASTOMTE = xDBIASTO0.DBIASTOMTE Then
                            newDBIASTO0.DBIASTOMTE = oldDBIASTO0.DBIASTOMTE
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDBIASTO0.DBIASTOMTE = xDBIASTO0.DBIASTOMTE
                            fgDetail = Format$(newDBIASTO0.DBIASTOMTE, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
        Case "DBIASTOAUT": xDBIASTO0.DBIASTOAUT = txtDetail_A
                        If oldDBIASTO0.DBIASTOAUT = xDBIASTO0.DBIASTOAUT Then
                           newDBIASTO0.DBIASTOAUT = oldDBIASTO0.DBIASTOAUT
                           fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDBIASTO0.DBIASTOAUT = xDBIASTO0.DBIASTOAUT
                            fgDetail = newDBIASTO0.DBIASTOAUT: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
                        
          Case "DBIASTOAU0": xDBIASTO0.DBIASTOAU0 = txtDetail_A
                        If oldDBIASTO0.DBIASTOAU0 = xDBIASTO0.DBIASTOAU0 Then
                           newDBIASTO0.DBIASTOAU0 = oldDBIASTO0.DBIASTOAU0
                           fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDBIASTO0.DBIASTOAU0 = xDBIASTO0.DBIASTOAU0
                            fgDetail = newDBIASTO0.DBIASTOAU0: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "DBIASTOCPT": xDBIASTO0.DBIASTOCPT = txtDetail_A
                        If oldDBIASTO0.DBIASTOCPT = xDBIASTO0.DBIASTOCPT Then
                           newDBIASTO0.DBIASTOCPT = oldDBIASTO0.DBIASTOCPT
                           fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDBIASTO0.DBIASTOCPT = xDBIASTO0.DBIASTOCPT
                            fgDetail = newDBIASTO0.DBIASTOCPT: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
   Case "YSTOPCI": xDBIASTO0.YSTOPCI = txtDetail_A
                        If oldDBIASTO0.YSTOPCI = xDBIASTO0.YSTOPCI Then
                           newDBIASTO0.YSTOPCI = oldDBIASTO0.YSTOPCI
                           fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDBIASTO0.YSTOPCI = xDBIASTO0.YSTOPCI
                            fgDetail = newDBIASTO0.YSTOPCI: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "YSTODEV": xDBIASTO0.YSTODEV = txtDetail_A
                        If oldDBIASTO0.YSTODEV = xDBIASTO0.YSTODEV Then
                            newDBIASTO0.YSTODEV = oldDBIASTO0.YSTODEV
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                             X = fraDetail_Control_YSTODEV(xDBIASTO0.YSTODEV)
                            If X = "" Then
                                Call MsgBox("devise inconnue : " & xDBIASTO0.YSTODEV, vbExclamation, "DBIASTO0 contrôle devise")
                            Else
                                Call lstErr_Clear(lstErr, cmdContext, xDBIASTO0.YSTODEV & " : " & X): DoEvents
                                newDBIASTO0.YSTODEV = xDBIASTO0.YSTODEV
                                fgDetail = newDBIASTO0.YSTODEV: fgDetail.CellBackColor = txtDetail_ColorUpdate
                            End If
                        End If
       Case "YSTOMON": xDBIASTO0.YSTOMON = txtDetail_C
                        If oldDBIASTO0.YSTOMON = xDBIASTO0.YSTOMON Then
                            newDBIASTO0.YSTOMON = oldDBIASTO0.YSTOMON
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDBIASTO0.YSTOMON = xDBIASTO0.YSTOMON
                            fgDetail = Format$(newDBIASTO0.YSTOMON, "### ### ### ###.00"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
            Case "YSTODEB": xDBIASTO0.YSTODEB = txtDetail_N
                        If oldDBIASTO0.YSTODEB = xDBIASTO0.YSTODEB Then
                           newDBIASTO0.YSTODEB = oldDBIASTO0.YSTODEB
                           fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            X = dateElp("DateMS", 0, xDBIASTO0.YSTODEB)
                            If X = "" Then
                                Call MsgBox("date non conforme 'aaaammjj' : " & xDBIASTO0.YSTODEB, vbExclamation, "DGAPPIS0 contrôle date")
                            Else
                               newDBIASTO0.YSTODEB = xDBIASTO0.YSTODEB
                                 fgDetail = newDBIASTO0.YSTODEB: fgDetail.CellBackColor = txtDetail_ColorUpdate
                             End If
                        End If
            Case "YSTOFIN": xDBIASTO0.YSTOFIN = txtDetail_N
                         If oldDBIASTO0.YSTOFIN = xDBIASTO0.YSTOFIN Then
                            newDBIASTO0.YSTOFIN = oldDBIASTO0.YSTOFIN
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                         Else
                            X = dateElp("DateMS", 0, xDBIASTO0.YSTOFIN)
                            If X = "" Then
                                Call MsgBox("date non conforme 'aaaammjj' : " & xDBIASTO0.YSTOFIN, vbExclamation, "DGAPPIS0 contrôle date")
                            Else
                                newDBIASTO0.YSTOFIN = xDBIASTO0.YSTOFIN
                                fgDetail = newDBIASTO0.YSTOFIN: fgDetail.CellBackColor = txtDetail_ColorUpdate
                            End If
                        End If
         Case "YSTOAPP": xDBIASTO0.YSTOAPP = txtDetail_A
                        If oldDBIASTO0.YSTOAPP = xDBIASTO0.YSTOAPP Then
                            newDBIASTO0.YSTOAPP = oldDBIASTO0.YSTOAPP
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDBIASTO0.YSTOAPP = xDBIASTO0.YSTOAPP
                            fgDetail = newDBIASTO0.YSTOAPP: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
        Case "YSTONAT": xDBIASTO0.YSTONAT = txtDetail_A
                        If oldDBIASTO0.YSTONAT = xDBIASTO0.YSTONAT Then
                            newDBIASTO0.YSTONAT = oldDBIASTO0.YSTONAT
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDBIASTO0.YSTONAT = xDBIASTO0.YSTONAT
                            fgDetail = newDBIASTO0.YSTONAT: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
        Case "YSTOCC1": xDBIASTO0.YSTOCC1 = txtDetail_A
                             If oldDBIASTO0.YSTOCC1 = xDBIASTO0.YSTOCC1 Then
                                newDBIASTO0.YSTOCC1 = oldDBIASTO0.YSTOCC1
                                fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                             Else
                                 newDBIASTO0.YSTOCC1 = xDBIASTO0.YSTOCC1
                                 fgDetail = newDBIASTO0.YSTOCC1: fgDetail.CellBackColor = txtDetail_ColorUpdate
                             End If
        Case "YSTOCL1": xDBIASTO0.YSTOCL1 = txtDetail_N
                             If oldDBIASTO0.YSTOCL1 = xDBIASTO0.YSTOCL1 Then
                                newDBIASTO0.YSTOCL1 = oldDBIASTO0.YSTOCL1
                                fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                             Else
                                 newDBIASTO0.YSTOCL1 = xDBIASTO0.YSTOCL1
                                 fgDetail = newDBIASTO0.YSTOCL1: fgDetail.CellBackColor = txtDetail_ColorUpdate
                             End If
         Case "YSTOCC2": xDBIASTO0.YSTOCC2 = txtDetail_A
                             If oldDBIASTO0.YSTOCC2 = xDBIASTO0.YSTOCC2 Then
                                newDBIASTO0.YSTOCC2 = oldDBIASTO0.YSTOCC2
                                fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                             Else
                                 newDBIASTO0.YSTOCC2 = xDBIASTO0.YSTOCC2
                                 fgDetail = newDBIASTO0.YSTOCC2: fgDetail.CellBackColor = txtDetail_ColorUpdate
                             End If
        Case "YSTOCL2": xDBIASTO0.YSTOCL2 = txtDetail_N
                             If oldDBIASTO0.YSTOCL2 = xDBIASTO0.YSTOCL2 Then
                                newDBIASTO0.YSTOCL2 = oldDBIASTO0.YSTOCL2
                                fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                             Else
                                 newDBIASTO0.YSTOCL2 = xDBIASTO0.YSTOCL2
                                 fgDetail = newDBIASTO0.YSTOCL2: fgDetail.CellBackColor = txtDetail_ColorUpdate
                             End If
        Case "YSTOCTX": xDBIASTO0.YSTOCTX = txtDetail_A
                        If oldDBIASTO0.YSTOCTX = xDBIASTO0.YSTOCTX Then
                            newDBIASTO0.YSTOCTX = oldDBIASTO0.YSTOCTX
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDBIASTO0.YSTOCTX = xDBIASTO0.YSTOCTX
                            fgDetail = newDBIASTO0.YSTOCTX: fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If
       Case "YSTOTAU": xDBIASTO0.YSTOTAU = txtDetail_D
                        If oldDBIASTO0.YSTOTAU = xDBIASTO0.YSTOTAU Then
                            newDBIASTO0.YSTOTAU = oldDBIASTO0.YSTOTAU
                            fgDetail = "": fgDetail.CellBackColor = &HFFFFFA
                        Else
                            newDBIASTO0.YSTOTAU = xDBIASTO0.YSTOTAU
                            fgDetail = Format$(newDBIASTO0.YSTOTAU, "### ###.000 000"): fgDetail.CellBackColor = txtDetail_ColorUpdate
                        End If

End Select
fgDetail.Row = fgDetail_RowClick
txtDetail_blnUpdate = False: cmdDetail_Control.Visible = False
cmdDetail_Update.Visible = DBIASTO0_Aut.Saisir
End Sub

Public Function fraDetail_Control_DBIASTOCLI(lDBIASTOCLI As Long)
Dim xSql As String
fraDetail_Control_DBIASTOCLI = ""
xSql = "select CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & Format$(lDBIASTOCLI, "0000000") & "'"
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then fraDetail_Control_DBIASTOCLI = rsSab("CLIENARA1")

End Function

Public Function fraDetail_Control_YSTODEV(lYSTODEV As String)
Dim xSql As String
fraDetail_Control_YSTODEV = ""
xSql = "select BASDVSABR from " & paramIBM_Library_SAB & ".ZBASDVS0 where BASDVSDEV = '" & lYSTODEV & "'"
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then fraDetail_Control_YSTODEV = rsSab("BASDVSABR")

End Function

Private Sub txtSelect_YSTOAPP_Click()
cmdSelect_Reset

End Sub

Private Sub txtSelect_YSTOAPP_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub

Private Sub txtSelect_YSTOPCI_Change()
cmdSelect_Reset

End Sub


Private Sub txtSelect_YSTOPCI_Click()
cmdSelect_Reset

End Sub


Private Sub txtSelect_YSTOPCI_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


