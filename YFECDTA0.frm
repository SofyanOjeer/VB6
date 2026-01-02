VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmYFECDTA0 
   AutoRedraw      =   -1  'True
   Caption         =   "FEC: fichier des écritures comptables"
   ClientHeight    =   12165
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   16335
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "YFECDTA0.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   12165
   ScaleWidth      =   16335
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   8685
      TabIndex        =   2
      Top             =   15
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   11640
      Left            =   30
      TabIndex        =   3
      Top             =   435
      Width           =   16290
      _ExtentX        =   28734
      _ExtentY        =   20532
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Rechercher"
      TabPicture(0)   =   "YFECDTA0.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "YFECDTA0.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFg"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "YFECDTA0.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraUpdate"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraUpdate 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6000
         Left            =   -66000
         TabIndex        =   14
         Top             =   1170
         Visible         =   0   'False
         Width           =   6015
         Begin VB.TextBox txtUpdate_CRETXTINFO 
            BackColor       =   &H00D0FFD0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3945
            Left            =   75
            MaxLength       =   1024
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   855
            Width           =   5790
         End
         Begin VB.CommandButton cmdUpdate_Quit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Abandonner"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   165
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   5070
            Width           =   1200
         End
         Begin VB.CommandButton cmdUpdate_Ok 
            BackColor       =   &H0000FF00&
            Caption         =   "Enregistrer"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   4380
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   5040
            Width           =   1230
         End
         Begin VB.Label lblUpdate_CREANOLTXT 
            BackColor       =   &H00C0E0FF&
            Caption         =   "saisir un commentaire :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   330
            Left            =   1455
            TabIndex        =   18
            Top             =   375
            Width           =   2700
         End
      End
      Begin VB.TextBox txtFg 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   -69030
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Text            =   "YFECDTA0.frx":035E
         Top             =   1155
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Frame fraSelect 
         BackColor       =   &H00E0E0E0&
         Height          =   11055
         Left            =   60
         TabIndex        =   4
         Top             =   540
         Width           =   16155
         Begin RichTextLib.RichTextBox txtRTF 
            Height          =   3735
            Left            =   960
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   7185
            Visible         =   0   'False
            Width           =   14775
            _ExtentX        =   26061
            _ExtentY        =   6588
            _Version        =   393217
            BackColor       =   14745599
            HideSelection   =   0   'False
            ScrollBars      =   3
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"YFECDTA0.frx":0366
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
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   9675
            Left            =   10365
            TabIndex        =   12
            Top             =   1320
            Visible         =   0   'False
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   17066
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   400
            BackColor       =   15790320
            ForeColor       =   4210752
            BackColorFixed  =   8421504
            ForeColorFixed  =   16777215
            BackColorBkg    =   15790320
            GridColor       =   10526720
            GridColorFixed  =   10526720
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            FormatString    =   "<? |<Opération                              |< Date                 |> N° CRE          |<Doc |<Comment|>nb    |"
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
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   13455
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   630
            Width           =   1335
         End
         Begin VB.ComboBox cboSelect_SQL 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   11610
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   180
            Width           =   4155
         End
         Begin VB.Frame fraSelect_Options 
            BackColor       =   &H00F0FFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1140
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Visible         =   0   'False
            Width           =   11205
            Begin VB.TextBox txtSelect_FECMVTSEQ 
               Height          =   285
               Left            =   5325
               TabIndex        =   20
               Top             =   465
               Width           =   1575
            End
            Begin VB.ComboBox cboSelect_FECLOGAA 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1185
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   435
               Width           =   1860
            End
            Begin VB.Label lblSelect_FECMVTSEQ 
               BackColor       =   &H00F0FFFF&
               Caption         =   "FEC séquence"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   4035
               TabIndex        =   19
               Top             =   480
               Width           =   1110
            End
            Begin VB.Label lblSelect_FECLOGAA 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Code état"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   225
               TabIndex        =   10
               Top             =   480
               Width           =   1155
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   9750
            Left            =   150
            TabIndex        =   9
            Top             =   1260
            Width           =   15825
            _ExtentX        =   27914
            _ExtentY        =   17198
            _Version        =   393216
            Cols            =   9
            FixedCols       =   0
            RowHeightMin    =   400
            BackColor       =   16777215
            ForeColor       =   16711680
            BackColorFixed  =   8421376
            ForeColorFixed  =   16777215
            BackColorBkg    =   16777215
            GridColor       =   10526720
            GridColorFixed  =   10526720
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            FormatString    =   $"YFECDTA0.frx":03E6
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
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
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
      Left            =   15690
      Picture         =   "YFECDTA0.frx":050F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   -30
      Width           =   500
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
      Begin VB.Menu mnuPrint_Excel 
         Caption         =   "Excel"
      End
      Begin VB.Menu mnuPrint_Mail 
         Caption         =   "Mail"
      End
   End
End
Attribute VB_Name = "frmYFECDTA0"
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
Dim arrHab(19) As Boolean
Dim blnAuto As Boolean, blnError As Boolean
Dim cmdSelect_SQL_K As String
Dim rsSab_X As New ADODB.Recordset
Dim mMail_Destinataires As String

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

Dim xYFECLOG0 As typeYFECLOG0, oldYFECLOG0 As typeYFECLOG0, newYFECLOG0 As typeYFECLOG0
Dim xYFECMVT0 As typeYFECMVT0, oldYFECMVT0 As typeYFECMVT0, newYFECMVT0 As typeYFECMVT0
Dim HeightOfLine As Long, LinesOfText As Long

Dim txtRTF_prtForeColor_Header As Long


Dim VB_RTF_Modèle As String


Dim T(100) As typeYFEC0, T_Nb As Integer
Dim devSD0 As Currency, devDB As Currency, devCR As Currency, devSD1 As Currency
Dim blnExportation_Ok As Boolean

Dim xYBIAMVTH As typeYBIAMVT0
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

Private Sub fgDetail_Display()
Dim X As String, xWhere As String
Dim xSQL As String

On Error GoTo Error_Handler

currentAction = "fgDetail_Display"
fgDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString
fgDetail.Row = 0
'___________________________________________________________________________

  
Do While Not rsSab.EOF

    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    'Call rsYFECLOG0_GetBuffer(rsSab, xYFECLOG0)
    fgDetail_Display_Line
    
    rsSab.MoveNext

Loop

fgDetail.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgDetail.Rows - 1): DoEvents

'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub



Public Sub fgDetail_Display_Line()
On Error Resume Next

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



Public Sub Form_Init()
Dim V, xSQL As String, X As String
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

fgDetail_FormatString = fgDetail.FormatString
fgDetail.Enabled = True
fgDetail.Visible = False
fgDetail.Top = fgSelect.Top
fgDetail.Left = fgSelect.Left + fgSelect.Width - fgDetail.Width - 200

fraSelect_Options.Visible = True


If cboSelect_SQL.ListCount > 0 Then cboSelect_SQL.ListIndex = 0
cboSelect_FECLOGAA.Clear
cboSelect_FECLOGAA.AddItem " "

xSQL = "select distinct FECLOGAA from " & paramIBM_Library_SABSPE & ".YFECLOG0 " _
     & " order by feclogaa desc" _

Set rsSab_X = cnsab.Execute(xSQL)
Do Until rsSab_X.EOF
    cboSelect_FECLOGAA.AddItem rsSab_X("FECLOGAA")
    rsSab_X.MoveNext
Loop
If cboSelect_FECLOGAA.ListCount > 1 Then
    cboSelect_FECLOGAA.ListIndex = 1
Else
    cboSelect_FECLOGAA.ListIndex = 0
End If
blnControl = True

    '
txtRTF.LoadFile paramServer("\\BiaDoc\Filigrane\VB_RTF_Modèle.rtf")
VB_RTF_Modèle = txtRTF.TextRTF

txtRTF.Visible = False

Set fraUpdate.Container = fraSelect
fraUpdate.Top = fgSelect.Top
fraUpdate.Left = fgSelect.Left + fgSelect.Width - fraUpdate.Width - 200

cmdSelect_Reset

Call cmdSelect_SQL_1
Me.Enabled = True

End Sub



'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
currentActiveControl_Name = C.Name
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
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

'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub


Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String, wIndex As Long

For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = lK
    Select Case lK
'        Case 3: fgSelect.Col = 3: X = Format$(Val(fgSelect.Text), "000000000000000.00")

    End Select
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I

fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
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



Private Sub fgSelect_Display_1()

Dim K As Long

On Error GoTo Error_Handler
currentAction = "fgSelect_Display_1"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = fgSelect_FormatString

fgSelect.Rows = 1
                 
fgSelect.Row = 0

Do While Not rsSab.EOF

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_Display_1_Line
    
    rsSab.MoveNext

Loop

fgSelect.Visible = True

'If fgSelect.Rows = 2 Then
'    fgSelect.Col = 0
'    Call fgDetail_Display(Trim(fgSelect.Text))
'End If

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgSelect_Display_3()

Dim K As Long

On Error GoTo Error_Handler
currentAction = "fgSelect_Display_3"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = "< Date TRT          |<Référence                                     |< Compte                      |>Débit                      |>Crédit                    |<Libellé                                                                                                                         |                       |"

fgSelect.Rows = 1
                 
fgSelect.Row = 0

Do While Not rsSab.EOF

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    Call rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVTH)
    fgSelect_Display_3_Line
    
    rsSab.MoveNext

Loop

fgSelect.Visible = True

'If fgSelect.Rows = 2 Then
'    fgSelect.Col = 0
'    Call fgDetail_Display(Trim(fgSelect.Text))
'End If

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgSelect_Display_3_SD0()

Dim K As Long, xCur As Currency
Dim xSQL As String

On Error GoTo Error_Handler
currentAction = "fgSelect_Display_3"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = "< Compte                                   |< Intitulé                                                       |>Solde début exercice                   |>Débit exercice                          |>Crédit exercice                |>Solde fin exercice                    |"

fgSelect.Rows = 1
                 
fgSelect.Row = 0

If Not rsSab.EOF Then

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    
    fgSelect.Col = 0: fgSelect.Text = rsSab("FECCPTCOM")
    fgSelect.Col = 2
    xCur = rsSab("FECCPTSD0")
    fgSelect.Text = Format$(Abs(xCur), "### ### ### ##0.00")
    If xCur > 0 Then
        fgSelect.CellForeColor = vbRed
    Else
        fgSelect.CellForeColor = vbBlue
    End If
    
    fgSelect.Col = 3: fgSelect.Text = Format$(rsSab("FECCPTDB"), "### ### ### ##0.00"): fgSelect.CellForeColor = vbRed
    fgSelect.Col = 4: fgSelect.Text = Format$(rsSab("FECCPTCR"), "### ### ### ##0.00")
    
    fgSelect.Col = 5
    xCur = rsSab("FECCPTSD1")
    fgSelect.Text = Format$(Abs(xCur), "### ### ### ##0.00")
    If xCur > 0 Then
        fgSelect.CellForeColor = vbRed
    Else
        fgSelect.CellForeColor = vbBlue
    End If


    xSQL = "select COMPTEINT from " & paramIBM_Library_SAB & ".ZCOMPTE0 " _
         & " where COMPTEETA = 1 and COMPTEPLA = 1 and COMPTECOM = '" & xYFECMVT0.FECMVTCOM & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then fgSelect.Col = 1: fgSelect.Text = rsSab("COMPTEINT")

End If

fgSelect.Visible = True


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub fgSelect_Display_2()

Dim K As Long, I As Long, mDEV As String

On Error GoTo Error_Handler
currentAction = "fgSelect_Display_2"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = "<Devise|<Classe|> Solde 01-01                                |> Débit                                       |> Crédit                                            |> Solde 31-12                                             "
fgSelect.Rows = 1
                 
fgSelect.Row = 0
fgSelect.Col = 2: fgSelect.CellAlignment = 1
fgSelect.Col = 3: fgSelect.CellAlignment = 1
fgSelect.Col = 4: fgSelect.CellAlignment = 1
fgSelect.Col = 5: fgSelect.CellAlignment = 1
mDEV = "": devSD0 = 0: devDB = 0: devCR = 0: devSD1 = 0

For K = 1 To T_Nb

    For I = 1 To 9
        If mDEV <> T(K).FEC_DEV Then Call fgSelect_Display_2_Total(mDEV): mDEV = T(K).FEC_DEV
        
        If T(K).FEC_SD0(I) = 0 And T(K).FEC_DB(I) = 0 And T(K).FEC_CR(I) = 0 Then
        Else
            
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            fgSelect.Col = 0: fgSelect.Text = T(K).FEC_DEV
            fgSelect.Col = 1: fgSelect.Text = I
            If T(K).FEC_SD0(I) <> 0 Then
                fgSelect.Col = 2: fgSelect.Text = Format(T(K).FEC_SD0(I), "### ### ### ### ##0.00")
                If T(K).FEC_SD0(I) < 0 Then fgSelect.CellForeColor = vbRed
            End If
            
            If T(K).FEC_DB(I) <> 0 Then fgSelect.Col = 3: fgSelect.Text = Format(T(K).FEC_DB(I), "### ### ### ### ##0.00"): fgSelect.CellForeColor = vbRed
            If T(K).FEC_CR(I) <> 0 Then fgSelect.Col = 4: fgSelect.Text = Format(T(K).FEC_CR(I), "### ### ### ### ##0.00")
            T(K).FEC_SD1(I) = T(K).FEC_SD0(I) - T(K).FEC_DB(I) + T(K).FEC_CR(I)
            If T(K).FEC_SD1(I) <> 0 Then
                fgSelect.Col = 5: fgSelect.Text = Format(T(K).FEC_SD1(I), "### ### ### ### ##0.00")
                If T(K).FEC_SD1(I) < 0 Then fgSelect.CellForeColor = vbRed
            End If
            devSD0 = devSD0 + T(K).FEC_SD0(I)
            devDB = devDB + T(K).FEC_DB(I)
            devCR = devCR + T(K).FEC_CR(I)
            devSD1 = devSD1 + T(K).FEC_SD1(I)
        End If
    Next I
    
Next K
Call fgSelect_Display_2_Total(mDEV)
fgSelect.Visible = True


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub fgSelect_Display_1_Line()
Dim K As Integer, wColor As Long

On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = dateImp10_S(rsSab("FECLOGAMJ")) & " " & timeImp8(rsSab("FECLOGHMS")) & " " & rsSab("FECLOGSEQ")
fgSelect.Col = 1: fgSelect.Text = rsSab("FECLOGUSR")
fgSelect.Col = 2: fgSelect.Text = rsSab("FECLOGK")
fgSelect.Col = 3: fgSelect.Text = rsSab("FECLOGAA"): fgSelect.CellFontBold = True
fgSelect.Col = 4: fgSelect.Text = rsSab("FECLOGSTA")
If rsSab("FECLOGNB") <> 0 Then fgSelect.Col = 5: fgSelect.Text = Format(rsSab("FECLOGNB"), "### ### ###")
fgSelect.Col = 6: fgSelect.Text = rsSab("FECLOGTXT")



If InStr(rsSab("FECLOGTXT"), "GÉNÉRÉS") Then
    fgSelect.CellBackColor = mColor_G1
Else
    If InStr(rsSab("FECLOGTXT"), "SUPPRI") Then
        fgSelect.CellBackColor = mColor_W0
    Else
        If InStr(rsSab("FECLOGTXT"), "INITIA") Then
            fgSelect.CellBackColor = mColor_Y1
        Else
            If InStr(rsSab("FECLOGTXT"), "Exportation début") Then
                fgSelect.CellBackColor = mColor_Y2
            Else
                If InStr(rsSab("FECLOGTXT"), "Exportation terminée") Then
                    fgSelect.CellBackColor = mColor_G2
                End If
            End If
        End If
        '
    End If
End If
If rsSab("FECLOGSTA") <> " " Then
    For K = 0 To 6
        fgSelect.Col = K
        fgSelect.CellBackColor = vbRed
        fgSelect.CellForeColor = vbYellow
    Next K
End If

End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim wFct As String

mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

wFct = UCase$(Trim(Mid$(Msg, 1, 12)))
Call BIA_VB_HAB(wFct, arrHab(), cboSelect_SQL)

Form_Init

mMail_Destinataires = currentSSIWINMAIL

Select Case wFct
    Case Else: blnAuto = False:

End Select
End Sub



Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgSelect.Visible = False
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
    fgSelect.Row = lRow
    For I = 1 To 0 Step -1
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = 1 To 0 Step -1
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
    End If
End If
fgSelect.LeftCol = fgSelect.FixedCols
fgSelect.Visible = True
End Sub


Private Sub cboSelect_FECLOGAA_Click()
cmdSelect_Clear
End Sub

Private Sub cboSelect_SQL_Click()
cmdSelect_Reset

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

Private Sub cmdSAB_Dossier_DB_Click()
End Sub

Private Sub cmdUpdate_Ok_Click()
Dim xSQL As String, xSet As String, X As String
On Error GoTo Error_Handler

GoTo Exit_sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdUpdate_Quit_Click()
fraUpdate.Visible = False

End Sub

Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, K As Integer
On Error Resume Next
txtRTF.Visible = False
fraUpdate.Visible = False


If y <= fgDetail.RowHeightMin Then
    fgDetail.Visible = False
    Select Case fgDetail.Col
        Case 0: fgDetail_Sort1 = 0: fgDetail_Sort2 = 3: fgDetail_Sort
        Case 1:  fgDetail_Sort1 = 1: fgDetail_Sort2 = 3: fgDetail_Sort
        Case 2: fgDetail_Sort1 = 2: fgDetail_Sort2 = 2: fgDetail_Sort
        Case 3: fgDetail_Sort1 = 3: fgDetail_Sort2 = 3: fgDetail_Sort
        Case 4: fgDetail_Sort1 = 4: fgDetail_Sort2 = 4: fgDetail_Sort
    End Select
    fgDetail.Visible = True
Else
    If fgDetail.Rows > 1 Then
        Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
        Select Case cmdSelect_SQL_K
            Case "1"
                
            Case "2"
 '               fgDetail.Col = 1: wX = Trim(fgDetail.Text)
        End Select
    End If
End If
fgDetail.LeftCol = 0



End Sub




Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, K As Integer, xSQL As String
On Error Resume Next

txtRTF.Visible = False
fraUpdate.Visible = False
If y <= fgSelect.RowHeightMin Then
    fgSelect.Visible = False
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
    End Select
    fgSelect.Visible = True
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        Select Case cmdSelect_SQL_K
            Case "1"
            Case "2"
                fgSelect.Col = 1: wX = Trim(fgSelect.Text)
            Case "3":
                fgSelect.Col = 2: wX = Trim(fgSelect.Text)
                Call frmSAB_Dossier_DB.Form_Init("MOUVEMDTR", wX, xYFECMVT0.FECMVTAA & "0101", xYFECMVT0.FECMVTAA & "1231", "", "", "", 0)

        End Select
        
   End If
End If
fgSelect.LeftCol = 0


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

Public Sub cmdSelect_Clear()

lstErr.Clear
fgSelect.Visible = False
fgDetail.Visible = False
txtRTF.Visible = False
cmdSelect_Ok.BackColor = vbGreen
fraUpdate.Visible = False
End Sub

Public Sub cmdSelect_Reset()
Dim K As Integer
If blnControl Then
    cmdSelect_Clear
    K = InStr(cboSelect_SQL, "-")
    If K > 1 Then
        cmdSelect_SQL_K = Trim(Mid$(cboSelect_SQL, 1, K - 1))
    Else
        cmdSelect_SQL_K = "???"
    End If
    
    fraSelect_Options.Visible = False
     txtSelect_FECMVTSEQ.Visible = False
    Select Case cmdSelect_SQL_K
        Case "1": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
        Case "3": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True: txtSelect_FECMVTSEQ.Visible = True
    End Select

End If
End Sub


Private Sub cmdSelect_SQL_3()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_3"
xWhere = ""

  
If Trim(cboSelect_FECLOGAA) = "" Then
    V = "Préciser l'exercice"
    GoTo Error_MsgBox
End If
If Val(txtSelect_FECMVTSEQ) = 0 Then
    V = "Préciser le numéro de séquence du fichier FEC"
    GoTo Error_MsgBox
End If
xYFECMVT0.FECMVTAA = Val(cboSelect_FECLOGAA)
xYFECMVT0.FECMVTSEQ = Val(txtSelect_FECMVTSEQ)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YFECMVT0 " _
     & " where FECMVTAA = " & xYFECMVT0.FECMVTAA & " and FECMVTSEQ = " & xYFECMVT0.FECMVTSEQ
Set rsSab = cnsab.Execute(xSQL)
  
If rsSab.EOF Then
    V = "numéro de séquence du fichier FEC INCONNU"
    GoTo Error_MsgBox
End If
xYFECMVT0.FECMVTPIE = rsSab("FECMVTPIE")
xYFECMVT0.FECMVTECR = rsSab("FECMVTECR")
xYFECMVT0.FECMVTCOM = rsSab("FECMVTCOM")

If xYFECMVT0.FECMVTPIE > 0 Then
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTHP " _
         & " where MOUVEMETA = 1 and MOUVEMPIE = " & xYFECMVT0.FECMVTPIE  '& " and MOUVEMECR = " & xYFECMVT0.FECMVTECR
    Set rsSab = cnsab.Execute(xSQL)
    
    Call fgSelect_Display_3
Else
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YFECCPT0 " _
         & " where FECCPTAA = " & xYFECMVT0.FECMVTAA & " and FECCPTCOM = '" & xYFECMVT0.FECMVTCOM & "'"
    Set rsSab = cnsab.Execute(xSQL)
    
    Call fgSelect_Display_3_SD0
End If


Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub fgSelect_Display_3_Line()
Dim K As Integer
'Dim wColor As Long, wColor_Row As Long
Dim X As String
On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = "  " & dateImp10_S(xYBIAMVTH.MOUVEMDTR + 19000000)
fgSelect.Col = 1:
    fgSelect.Text = xYBIAMVTH.MOUVEMSER & " " & xYBIAMVTH.MOUVEMSSE & " " & xYBIAMVTH.MOUVEMOPE & " " & xYBIAMVTH.MOUVEMNUM & " " & xYBIAMVTH.MOUVEMEVE

fgSelect.Col = 2: fgSelect.Text = xYBIAMVTH.MOUVEMCOM

fgSelect.Col = IIf(xYBIAMVTH.MOUVEMMON > 0, 3, 4)

fgSelect.Text = Format$(Abs(xYBIAMVTH.MOUVEMMON), "### ### ### ##0.00")

If xYBIAMVTH.MOUVEMMON > 0 Then
    fgSelect.CellForeColor = vbRed
Else
    fgSelect.CellForeColor = vbBlue
End If

fgSelect.Col = 5: fgSelect.Text = Trim(xYBIAMVTH.LIBELLIB1) & Trim(xYBIAMVTH.LIBELLIB2) & Trim(xYBIAMVTH.LIBELLIB3) & Trim(xYBIAMVTH.LIBELLIB4)
fgSelect.Col = 6:
X = Format$(xYBIAMVTH.MOUVEMPIE, "##### ##0") & "-" & Format$(xYBIAMVTH.MOUVEMECR, "### ##0")
fgSelect.Text = X
fgSelect.Col = fgSelect_arrIndex
    fgSelect.Text = xYBIAMVTH.MOUVEMDTR & X
If xYBIAMVTH.MOUVEMECR = xYFECMVT0.FECMVTECR Then
    For K = 0 To 6: fgSelect.Col = K: fgSelect.CellBackColor = mColor_Y2: Next K
End If
End Sub

Private Sub cmdSelect_SQL_1()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1"
xWhere = ""
If Trim(cboSelect_FECLOGAA) <> "" Then xWhere = " where FECLOGAA = " & Trim(cboSelect_FECLOGAA)
    
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YFECLOG0 " & xWhere & " order by FECLOGAMJ , FECLOGHMS , FECLOGSEQ"
Set rsSab = cnsab.Execute(xSQL)
  
Call fgSelect_Display_1


Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_2()
Dim V, X As String, wFile As String
Dim xSQL As String, xWhere As String
Dim intFile As Integer, Nb As Long
Dim kPCI As Integer, kDebit As Integer, kCredit As Integer, kDevise As Integer, kMTD As Integer, xDevise As String
Dim tDB As Currency, tCR As Currency, curDB As Currency, curCR As Currency, curX As Currency
Dim K As Integer, I As Integer
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2"
cmdSelect_Clear

If Trim(cboSelect_FECLOGAA) = "" Then
    V = "Préciser l'exercice"
    GoTo Error_MsgBox
End If
wFile = "c:\temp\" & socSiren & "FEC" & Trim(cboSelect_FECLOGAA) & "1231.txt"

Call rsYFECLOG0_Init(newYFECLOG0)
newYFECLOG0.FECLOGAA = Trim(cboSelect_FECLOGAA)
newYFECLOG0.FECLOGK = "YFECDTA0"
newYFECLOG0.FECLOGTXT = "Exportation début" & wFile
blnExportation_Ok = True

'V = cnSAB_Transaction("BeginTrans")
cnSab_Update.Open paramODBC_DSN_SAB
Call sqlYFECLOG0_Insert(newYFECLOG0)
cnSab_Update.Close
'V = cnSAB_Transaction("Commit")

T_Nb = 0
xSQL = "select distinct COMPTEDEV from " & paramIBM_Library_SAB & ".ZCOMPTE0  order by COMPTEDEV"

Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    T_Nb = T_Nb + 1
    T(T_Nb).FEC_DEV = rsSab("COMPTEDEV")
    For I = 0 To 9
        T(T_Nb).FEC_SD0(I) = 0: T(T_Nb).FEC_DB(I) = 0: T(T_Nb).FEC_CR(I) = 0: T(T_Nb).FEC_SD1(I) = 0
    Next I
    
    rsSab.MoveNext
Loop


xWhere = " where FECDTAaa = " & newYFECLOG0.FECLOGAA
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YFECDTA0 " & xWhere & " order by FECDTASEQ"

Set rsSab = cnsab.Execute(xSQL)

intFile = FreeFile(0)
  
Open wFile For Output As #intFile
Print #intFile, "JournalCode |JournalLib |EcritureNum |EcritureDate |CompteNum |CompteLib|CompAuxNum |CompAuxLib |PieceRef |PieceDate |EcritureLib |Debit |Credit |EcritureLet |DateLet |ValidDate | Montantdevise |ldevise"; ""


Do While Not rsSab.EOF
    Nb = Nb + 1
    X = rsSab("FECDTATXT")
    If Nb = 1 Then
        
        K = InStr(1, X, "|") + 1: K = InStr(K, X, "|") + 1: K = InStr(K, X, "|") + 1
        kPCI = InStr(K, X, "|") + 1
        
        K = InStr(kPCI, X, "|") + 1: K = InStr(K, X, "|") + 1: K = InStr(K, X, "|") + 1
        K = InStr(K, X, "|") + 1: K = InStr(K, X, "|") + 1: K = InStr(K, X, "|") + 1
        
        kDebit = InStr(K, X, "|") + 1
        kCredit = InStr(kDebit, X, "|") + 1
        
        K = InStr(kCredit, X, "|") + 1: K = InStr(K, X, "|") + 1: K = InStr(K, X, "|") + 1
        kMTD = InStr(K, X, "|") + 1
        kDevise = InStr(kMTD, X, "|") + 1
   End If
   
  
  I = Val(Mid$(X, kPCI, 1))
  xDevise = Mid$(X, kDevise, 3)
  If xDevise = "   " Then
        xDevise = "EUR"
        curDB = CCur(Mid$(X, kDebit, 18))
        curCR = CCur(Mid$(X, kCredit, 18))
   Else
    curX = CCur(Mid$(X, kMTD, 19))
    If curX > 0 Then
        curCR = curX: curDB = 0
    Else
         curDB = -curX: curCR = 0
   End If
    
   End If
    
  For K = 1 To T_Nb
    If xDevise = T(K).FEC_DEV Then Exit For
  Next K
  
    If Mid$(X, 1, 3) = "OUV" Then
        T(K).FEC_SD0(I) = T(K).FEC_SD0(I) + curCR - curDB
    Else
         T(K).FEC_DB(I) = T(K).FEC_DB(I) + curDB
         T(K).FEC_CR(I) = T(K).FEC_CR(I) + curCR
    End If
    
    If Nb Mod 10000 = 0 Then
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "> " & Nb & " enregistrements exportés"): DoEvents
        ''''Exit Do
    End If

    Print #intFile, X
    
    rsSab.MoveNext

Loop
    
Call fgSelect_Display_2
Call mnuPrint_Mail_Click
Call mnuPrint_Excel_Click

If Not blnExportation_Ok Then
    newYFECLOG0.FECLOGSTA = "E"
    MsgBox ("erreur DB <> CR")
End If
newYFECLOG0.FECLOGTXT = "Exportation terminée " & wFile
newYFECLOG0.FECLOGNB = Nb
cnSab_Update.Open paramODBC_DSN_SAB
Call sqlYFECLOG0_Insert(newYFECLOG0)
cnSab_Update.Close
    
Call lstErr_ChangeLastItem(lstErr, cmdContext, "= " & Nb & " enregistrements exportés"): DoEvents
Set rsSab = Nothing
Close #intFile
'____________________________________________________________________________________________________

'____________________________________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Return()
    If SSTab1.Tab = 0 Then
        If Not fraUpdate.Visible Then cmdSelect_Ok_Click
    Else
        SendKeys "{TAB}"
    End If
End Sub


Public Sub cmdContext_Quit()
lstErr.Clear: lstErr.Height = 200

If fraUpdate.Visible Then
    fraUpdate.Visible = False
    Exit Sub
End If

If txtRTF.Visible Then
    txtRTF.Visible = False
    Exit Sub
End If

If txtFg.Visible Then
    txtFg.Visible = False
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


Unload Me

End Sub

Private Sub Form_Load()


mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False

End Sub


Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Private Sub lstErr_Click()
If lstErr.Height > 500 Then
    lstErr.Height = 480
Else
    lstErr.Height = lstErr.ListCount * 200 + 300
End If

End Sub





Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> CPT_FEC_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Clear

Select Case cmdSelect_SQL_K
    Case "1": cmdSelect_SQL_1
    Case "2": cmdSelect_SQL_2
    Case "3": cmdSelect_SQL_3
'    Case "SPLF": cmdSelect_SQL_SPLF
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< CPT_FEC_cmdSelect_Ok"): DoEvents
lstErr.Height = 480
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
cmdSelect_Ok.BackColor = fgSelect.BackColorFixed
End Sub




Private Sub mnuPrint_Excel_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String
Call lstErr_AddItem(lstErr, cmdContext, "> CPT_FEC : export Excel ...."): DoEvents
    Select Case cmdSelect_SQL_K
        Case "1":
            X = "Historique des traitements FEC (legifrance.gouv.fr) " & dateImp10_S(DSys) & " " & Time
            Call MSflexGrid_Excel("", "CPT_FEC", X, fgSelect, 7)
        Case "2":
            X = "Balance exportation FEC (legifrance.gouv.fr) " & dateImp10_S(DSys) & " " & Time
            Call MSflexGrid_Excel("", "CPT_FEC", X, fgSelect, 7)
    End Select

Call lstErr_AddItem(lstErr, cmdContext, "< CPT_FEC : export Excel terminé"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint_Mail_Click()
Dim X As String

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_AddItem(lstErr, cmdContext, "> CPT_FEC : export mail ...."): DoEvents
    Select Case cmdSelect_SQL_K
        Case "1":
            X = "Historique des traitements FEC (legifrance.gouv.fr) au " & dateImp10_S(DSys) & " " & Time
            Call MSFlexGrid_SendMail(mMail_Destinataires, "CPT_FEC", X, X, fgSelect, 7)
        Case "2":
            X = "Balance exportation FEC (legifrance.gouv.fr) " & dateImp10_S(DSys) & " " & Time
            Call MSFlexGrid_SendMail(mMail_Destinataires, "CPT_FEC", X, X, fgSelect, 7)
    End Select

Call lstErr_AddItem(lstErr, cmdContext, "< CPT_FEC : export mail terminé"): DoEvents


Me.Enabled = True: Me.MousePointer = 0

End Sub






Public Sub fraDetail_Display()
Dim xSQL As String

fraDetail_Display_txtRTF
End Sub
Public Sub fraDetail_Display_txtRTF()
Dim xRTF As String, intFile As Integer, xIn As String

End Sub


Public Sub txtRTF_Visible()
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "{\f0\fswiss\fprq2\fcharset0 Calibri;}", "{\f0\fmodern\fprq1\fcharset0 Courier New;}")
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "\cf1\f0\fs20 a\cf2 b\cf3 c\cf4 d\cf5 e\cf6 f\cf7 g\cf8 h\cf9 i\cf10 j\cf11 k\cf12 l\cf13 m\cf14 n\cf15 o\cf16 p", "")
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", "")
txtRTF.Visible = True

End Sub



Public Sub fgSelect_Display_2_Total(lDEV As String)
On Error GoTo Error_Handler
Dim K As Integer
If devSD0 = 0 And devDB = 0 And devCR = 0 Then
Else
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect.Col = 0: fgSelect.Text = lDEV
    fgSelect.Col = 2: fgSelect.Text = Format(devSD0, "### ### ### ### ##0.00")
    fgSelect.Col = 3: fgSelect.Text = Format(devDB, "### ### ### ### ##0.00"): fgSelect.CellForeColor = vbRed
    fgSelect.Col = 4: fgSelect.Text = Format(devCR, "### ### ### ### ##0.00")
    fgSelect.Col = 5: fgSelect.Text = Format(devSD1, "### ### ### ### ##0.00")
    If devDB = devCR Then
        For K = 0 To 5: fgSelect.Col = K: fgSelect.CellBackColor = mColor_G1: Next K
    Else
        blnExportation_Ok = False
        For K = 0 To 5: fgSelect.Col = K: fgSelect.CellBackColor = mColor_W1: Next K
    End If
End If

devSD0 = 0: devDB = 0: devCR = 0: devSD1 = 0
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub

Private Sub txtSelect_FECMVTSEQ_Change()
cmdSelect_Clear

End Sub


Private Sub txtSelect_FECMVTSEQ_GotFocus()
Call txt_GotFocus(txtSelect_FECMVTSEQ)

End Sub


Private Sub txtSelect_FECMVTSEQ_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtSelect_FECMVTSEQ_LostFocus()
Call txt_LostFocus(txtSelect_FECMVTSEQ)

End Sub


