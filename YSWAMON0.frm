VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmYSWAMON0 
   AutoRedraw      =   -1  'True
   Caption         =   "SWAP_TAUX : MT360-362-364 "
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
   Icon            =   "YSWAMON0.frx":0000
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
      Top             =   420
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
      TabPicture(0)   =   "YSWAMON0.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "YSWAMON0.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "YSWAMON0.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtFg"
      Tab(2).ControlCount=   1
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
         Left            =   -71445
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Text            =   "YSWAMON0.frx":035E
         Top             =   2490
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Frame fraSelect 
         BackColor       =   &H00E0E0E0&
         Height          =   11055
         Left            =   75
         TabIndex        =   4
         Top             =   525
         Width           =   16155
         Begin VB.Frame fraSWI360 
            BackColor       =   &H00E0FFFF&
            Height          =   5655
            Left            =   8925
            TabIndex        =   15
            Top             =   1710
            Visible         =   0   'False
            Width           =   6585
            Begin VB.TextBox txtSWI360_22B 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   3315
               MaxLength       =   4
               TabIndex        =   27
               Top             =   3555
               Width           =   870
            End
            Begin VB.TextBox txtSWI360_14C 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   3285
               MaxLength       =   4
               TabIndex        =   25
               Top             =   2670
               Width           =   870
            End
            Begin VB.TextBox txtSWI360_CLI 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   350
               Left            =   1700
               MaxLength       =   10
               TabIndex        =   21
               Top             =   300
               Width           =   2040
            End
            Begin VB.CommandButton cmdSWI360_Delete 
               BackColor       =   &H00FF80FF&
               Caption         =   "Supprimer"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   4830
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   4530
               Width           =   900
            End
            Begin VB.CommandButton cmdSWI360_Add 
               BackColor       =   &H000080FF&
               Caption         =   "Ajouter"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   1905
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   4530
               Width           =   900
            End
            Begin VB.CommandButton cmdSWI360_Update 
               BackColor       =   &H0080FF80&
               Caption         =   "Enregistrer"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   3255
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   4530
               Width           =   900
            End
            Begin VB.CommandButton cmdSWI360_Quit 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Abandonner"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   420
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   4530
               Width           =   990
            End
            Begin VB.TextBox txtSWI360_77H 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   1500
               MaxLength       =   35
               TabIndex        =   16
               Top             =   2055
               Width           =   4035
            End
            Begin VB.Label libSWI360_CLI 
               BackColor       =   &H00E0FFFF&
               Caption         =   "libSWI360_CLI"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1725
               TabIndex        =   28
               Top             =   825
               Width           =   4560
            End
            Begin VB.Label lblSWI360_22B 
               BackColor       =   &H00E0FFFF&
               Caption         =   "22B : Financial Center   4!c"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   255
               TabIndex        =   26
               Top             =   3600
               Width           =   2685
            End
            Begin VB.Label lblSWI360_14C 
               BackColor       =   &H00E0FFFF&
               Caption         =   "14C : Year of Definition   4!n"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   255
               TabIndex        =   24
               Top             =   2715
               Width           =   2685
            End
            Begin VB.Label lblSWI360_CLI 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Racine"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   255
               TabIndex        =   23
               Top             =   345
               Width           =   1290
            End
            Begin VB.Label lblSWI360_77H 
               BackColor       =   &H00E0FFFF&
               Caption         =   "77H : Type, Date, Version of the Agreement   6a[/8!n][//4!n]"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   255
               TabIndex        =   22
               Top             =   1545
               Width           =   5640
            End
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
            Left            =   13905
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   645
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
            Left            =   135
            TabIndex        =   5
            Top             =   105
            Visible         =   0   'False
            Width           =   11205
            Begin VB.TextBox txtSelect_SWAMONNUM 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1905
               TabIndex        =   13
               Top             =   360
               Width           =   1575
            End
            Begin MSComCtl2.DTPicker txtSelect_SWAMONYAMJ_Min 
               Height          =   300
               Left            =   9315
               TabIndex        =   10
               Top             =   300
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
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
               CheckBox        =   -1  'True
               CustomFormat    =   "dd  MM yyy"
               Format          =   18415619
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_SWAMONYAMJ_Max 
               Height          =   300
               Left            =   9555
               TabIndex        =   11
               Top             =   705
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
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
               Format          =   18415619
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_SWAMONNUM 
               BackColor       =   &H00F0FFFF&
               Caption         =   "N° opération"
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
               Left            =   675
               TabIndex        =   12
               Top             =   390
               Width           =   1110
            End
            Begin VB.Label lblSelect_SWAMONYAMJ 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Période"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   8115
               TabIndex        =   9
               Top             =   480
               Width           =   855
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   9750
            Left            =   135
            TabIndex        =   8
            Top             =   1185
            Width           =   15825
            _ExtentX        =   27914
            _ExtentY        =   17198
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   400
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   8421376
            ForeColorFixed  =   16777215
            BackColorBkg    =   16777215
            GridColor       =   10526720
            GridColorFixed  =   10526720
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            FormatString    =   $"YSWAMON0.frx":0366
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
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
      Picture         =   "YSWAMON0.frx":0463
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
Attribute VB_Name = "frmYSWAMON0"
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

'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long

Dim xYSWAMON0 As typeYSWAMON0, oldYSWAMON0 As typeYSWAMON0, newYSWAMON0 As typeYSWAMON0
Dim HeightOfLine As Long, LinesOfText As Long


Dim mSWAMONZSWI As Long

Dim mCLIENACLI As String, wCLIENACLI As Long
Dim Old_YBIATAB0_77H As typeYBIATAB0, New_YBIATAB0 As typeYBIATAB0
Dim Old_YBIATAB0_14C As typeYBIATAB0, old_YBIATAB0_22B As typeYBIATAB0

Dim blnSWI360_77H As Boolean, blnSWI360_14C As Boolean, blnSWI360_22B As Boolean
Public Sub Form_Init()
Dim V, xSql As String, X As String
Dim K As Long

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True


cmdReset
blnControl = False
Call DTPicker_Set(txtSelect_SWAMONYAMJ_Min, YBIATAB0_DATE_CPT_JS1)
Call DTPicker_Set(txtSelect_SWAMONYAMJ_Max, YBIATAB0_DATE_CPT_JS1)
'txtSelect_SWAMONYAMJ_Min.value = Null

fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False
cboSelect_SQL.ListIndex = 0
fraSelect_Options.Visible = True

blnControl = True

cmdSelect_Reset
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

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents
fgSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgSelect_Display_9()

Dim K As Long

On Error GoTo Error_Handler
currentAction = "fgSelect_Display_9"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = "<Racine         |<Intitulé                                                          |<77H : Type,Date,Version of the agreement " _
                      & "|<14C : Year of definition|<22B Financial centre"
mCLIENACLI = ""
fgSelect.Rows = 1
                 
fgSelect.Row = 0
Do While Not rsSab.EOF
    
    fgSelect_Display_9_Line
    rsSab.MoveNext

Loop

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents
fgSelect.Visible = True
If fgSelect.Rows = 1 Then Call fraSWI360_Display(0)
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSelect_Display_1_Line()
Dim K As Integer, wColor As Long

On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = rsSab("SWAMONSER") & " " & rsSab("SWAMONSES") & " " & rsSab("SWAMONOPR") & " " & rsSab("SWAMONNUM") & "-" & rsSab("SWAMONHISV")
fgSelect.Col = 1: fgSelect.Text = rsSab("SWAMONNAT")
fgSelect.Col = 2: fgSelect.Text = rsSab("SWAMONMTK") & " " & rsSab("SWAMON22A")
fgSelect.Col = 3
Select Case Trim(rsSab("SWAMONSTAK"))
    Case "": wColor = mColor_Y1: fgSelect.Text = rsSab("SWAMONSTAK") & " - " & "swift à générer"
    Case "W": wColor = mColor_G0: fgSelect.Text = rsSab("SWAMONSTAK") & " - " & "swift généré"
    Case "E": wColor = mColor_W0: fgSelect.Text = rsSab("SWAMONSTAK") & " - " & "erreur de traitement"
    Case "H": wColor = RGB(245, 245, 245): fgSelect.Text = rsSab("SWAMONSTAK") & " - " & "historique sans génération swift"
    Case "?": wColor = mColor_W1: fgSelect.Text = rsSab("SWAMONSTAK") & " - " & "code message non programmé"
    Case Else: wColor = RGB(245, 245, 245): fgSelect.Text = rsSab("SWAMONSTAK") & " - " & "intervention manuelle"
End Select

If rsSab("SWAMONZSWI") <> 0 Then fgSelect.Col = 4: fgSelect.Text = rsSab("SWAMONZSWI")


fgSelect.Col = 5: fgSelect.Text = dateImp10_S(rsSab("SWAMONYAMJ") + 19000000) & " " & timeImp8(rsSab("SWAMONYHMS")) & "-" & rsSab("SWAMONYVER") & "  " & Trim(rsSab("SWAMONYUSR"))

    For K = 0 To 8
        fgSelect.Col = K
        fgSelect.CellBackColor = wColor
    Next K


End Sub

Public Sub fgSelect_Display_9_Line()
Dim K As Integer, wColor As Long

On Error Resume Next

If Trim(rsSab("BIATABK1")) <> mCLIENACLI Then
    mCLIENACLI = Trim(rsSab("BIATABK1"))
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1

    fgSelect.Col = 0: fgSelect.Text = mCLIENACLI
    fgSelect.Col = 1: fgSelect.Text = rsSab("CLIENARA1")
End If

Select Case Trim(rsSab("BIATABK2"))
    Case "77H": fgSelect.Col = 2
    Case "14C": fgSelect.Col = 3
    Case "22B": fgSelect.Col = 4
End Select
fgSelect.Text = Trim(rsSab("BIATABTXT"))
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

Private Sub cmdSWI360_Add_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Parametrage_SWI360_New
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdSWI360_Delete_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Parametrage_SWI360_Delete
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSWI360_Quit_Click()
fraSWI360.Visible = False
End Sub

Private Sub cmdSWI360_Update_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Parametrage_SWI360_Update
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, K As Integer, xSql As String
On Error Resume Next


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
                fgSelect.Col = 4: mSWAMONZSWI = Val(fgSelect.Text)
                ZSWIFTA0_Display
             Case "9"
                If arrHab(18) Then
                    fgSelect.Col = 0: wCLIENACLI = Val(fgSelect.Text)
                    Call fraSWI360_Display(wCLIENACLI)
                End If
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
fraSWI360.Visible = False
cmdSelect_Ok.BackColor = vbGreen
If Not IsNull(txtSelect_SWAMONYAMJ_Min.value) Then
    txtSelect_SWAMONYAMJ_Max.Visible = True
Else
    txtSelect_SWAMONYAMJ_Max.Visible = False
End If
End Sub

Private Sub txtSelect_SWAMONYAMJ_Max_Change()
cmdSelect_Clear

End Sub

Private Sub txtSelect_SWAMONYAMJ_Min_Change()
cmdSelect_Clear

End Sub

Private Sub txtSelect_SWAMONNUM_Change()
cmdSelect_Clear

End Sub


Private Sub txtSelect_SWAMONNUM_GotFocus()
Call txt_GotFocus(txtSelect_SWAMONNUM)

End Sub


Private Sub txtSelect_SWAMONNUM_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtSelect_SWAMONNUM_LostFocus()
Call txt_LostFocus(txtSelect_SWAMONNUM)
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
    Select Case cmdSelect_SQL_K
        Case "1": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
    End Select

End If
End Sub


Private Sub cmdSelect_SQL_1()
Dim V, X As String
Dim xSql As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1"
xWhere = ""
If Trim(txtSelect_SWAMONNUM) <> "" Then
    xWhere = " Where SWAMONNUM = " & Val(Trim(txtSelect_SWAMONNUM))
End If

If Not IsNull(txtSelect_SWAMONYAMJ_Min.value) Then
    Call DTPicker_Control(txtSelect_SWAMONYAMJ_Min, wAmjMin)
    Call DTPicker_Control(txtSelect_SWAMONYAMJ_Max, wAmjMax)
    If xWhere = "" Then
        xWhere = " where"
    Else
        xWhere = " and"
    End If
    xWhere = xWhere & " SWAMONYAMJ >= " & wAmjMin & " And SWAMONYAMJ <= " & wAmjMax
End If

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWAMON0 " & xWhere & " order by SWAMONSER , SWAMONSES , SWAMONOPR , SWAMONNUM , SWAMONYAMJ , SWAMONYHMS"
Set rsSab = cnsab.Execute(xSql)
  
Call fgSelect_Display_1


Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_9()
Dim V, X As String
Dim xSql As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1"
xWhere = ""

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0," & paramIBM_Library_SAB & ".ZCLIENA0" _
     & " where BIATABID = 'SWI360' and CLIENACLI = BIATABK1 order by  BIATABK1 , BIATABK2"
Set rsSab = cnsab.Execute(xSql)
  
Call fgSelect_Display_9


Set rsSab = Nothing

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
        cmdSelect_Ok_Click
    Else
        SendKeys "{TAB}"
    End If
End Sub


Public Sub cmdContext_Quit()
lstErr.Clear: lstErr.Height = 200

If fraSWI360.Visible Then
    fraSWI360.Visible = False
    Exit Sub
End If


If fgSelect.Visible Then
    fgSelect.Visible = False
    Exit Sub
End If


Unload Me

End Sub
Private Function Parametrage_SWI360_Delete()
Dim xSql As String
On Error GoTo Error_Handler

Dim V, V2
App_Debug = "Parametrage_SWI360_Delete"
New_YBIATAB0.BIATABID = "SWI360"
New_YBIATAB0.BIATABK1 = Format(Val(txtSWI360_CLI), "0000000")

If New_YBIATAB0.BIATABK1 <> Old_YBIATAB0_77H.BIATABK1 Then
    V = "La racine a été modifiée : " & Old_YBIATAB0_77H.BIATABK1 & " # " & New_YBIATAB0.BIATABK1
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
    GoTo END_Function
End If

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
If blnSWI360_77H Then V = sqlYBIATAB0_Delete(Old_YBIATAB0_77H)
If Not IsNull(V) Then GoTo Error_MsgBox
If blnSWI360_14C Then V = sqlYBIATAB0_Delete(Old_YBIATAB0_14C)
If Not IsNull(V) Then GoTo Error_MsgBox
If blnSWI360_22B Then V = sqlYBIATAB0_Delete(old_YBIATAB0_22B)
If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    Parametrage_SWI360_Delete = V
    If Not IsNull(V) Then
        V2 = cnSAB_Transaction("Rollback")
    Else
        V2 = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
If IsNull(V) Then
    fraSWI360.Visible = False
    Call cmdSelect_SQL_9
End If
END_Function:

End Function

Private Function Parametrage_SWI360_New()
Dim xSql As String
On Error GoTo Error_Handler

Dim V, V2
App_Debug = "Parametrage_SWI360_New"
New_YBIATAB0.BIATABID = "SWI360"
New_YBIATAB0.BIATABK1 = Format(Val(txtSWI360_CLI), "0000000")

xSql = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0" _
     & " where CLIENACLI = '" & New_YBIATAB0.BIATABK1 & "'"
Set rsSab = cnsab.Execute(xSql)
If rsSab.EOF Then
    V = "Racine inconnue : " & txtSWI360_CLI
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
    GoTo END_Function
End If

New_YBIATAB0.BIATABID = "SWI360"

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
New_YBIATAB0.BIATABK2 = "77H"
New_YBIATAB0.BIATABTXT = Trim(txtSWI360_77H)
V = sqlYBIATAB0_Insert(New_YBIATAB0)
If Not IsNull(V) Then GoTo Error_MsgBox

New_YBIATAB0.BIATABK2 = "14C"
New_YBIATAB0.BIATABTXT = Trim(txtSWI360_14C)
V = sqlYBIATAB0_Insert(New_YBIATAB0)
If Not IsNull(V) Then GoTo Error_MsgBox

New_YBIATAB0.BIATABK2 = "22B"
New_YBIATAB0.BIATABTXT = Trim(txtSWI360_22B)
V = sqlYBIATAB0_Insert(New_YBIATAB0)
If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    Parametrage_SWI360_New = V
    If Not IsNull(V) Then
        V2 = cnSAB_Transaction("Rollback")
    Else
        V2 = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
If IsNull(V) Then
    fraSWI360.Visible = False
    Call cmdSelect_SQL_9
End If
END_Function:
End Function


Public Function Parametrage_SWI360_Update()
Dim V, V2


App_Debug = "Parametrage_SWI360_New"
New_YBIATAB0.BIATABID = "SWI360"
New_YBIATAB0.BIATABK1 = Format(Val(txtSWI360_CLI), "0000000")

If New_YBIATAB0.BIATABK1 <> Old_YBIATAB0_77H.BIATABK1 Then
    V = "La racine a été modifiée : " & Old_YBIATAB0_77H.BIATABK1 & " # " & New_YBIATAB0.BIATABK1
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
    GoTo END_Function
End If
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
New_YBIATAB0.BIATABK2 = "77H"
New_YBIATAB0.BIATABTXT = Trim(txtSWI360_77H)
If blnSWI360_77H Then
    V = sqlYBIATAB0_Update(New_YBIATAB0, Old_YBIATAB0_77H)
Else
    V = sqlYBIATAB0_Insert(New_YBIATAB0)
End If
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
New_YBIATAB0.BIATABK2 = "14C"
New_YBIATAB0.BIATABTXT = Trim(txtSWI360_14C)
If blnSWI360_14C Then
    V = sqlYBIATAB0_Update(New_YBIATAB0, Old_YBIATAB0_14C)
Else
    V = sqlYBIATAB0_Insert(New_YBIATAB0)
End If
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
New_YBIATAB0.BIATABK2 = "22B"
New_YBIATAB0.BIATABTXT = Trim(txtSWI360_22B)
If blnSWI360_22B Then
    V = sqlYBIATAB0_Update(New_YBIATAB0, old_YBIATAB0_22B)
Else
    V = sqlYBIATAB0_Insert(New_YBIATAB0)
End If
If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    
    Parametrage_SWI360_Update = V

    If Not IsNull(V) Then
        V2 = cnSAB_Transaction("Rollback")
    Else
        V2 = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
If IsNull(V) Then
    fraSWI360.Visible = False
    Call cmdSelect_SQL_9
End If
END_Function:

End Function



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
Call lstErr_Clear(lstErr, cmdContext, "> SWAP_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Clear

Select Case cmdSelect_SQL_K
    Case "1": cmdSelect_SQL_1
    Case "9": cmdSelect_SQL_9
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< SWAP_cmdSelect_Ok"): DoEvents
lstErr.Height = 480
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
cmdSelect_Ok.BackColor = fgSelect.BackColorFixed
End Sub




Private Sub mnuPrint_Excel_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String
Call lstErr_AddItem(lstErr, cmdContext, "SWAP : export Excel ...."): DoEvents
    Select Case cmdSelect_SQL_K
        Case "1":
            X = "SWAP"
            Call MSflexGrid_Excel("", "SWAP", X, fgSelect, 7)
    End Select

Call lstErr_AddItem(lstErr, cmdContext, "Rétrocession de commissions  : export Excel terminé"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint_Mail_Click()
Dim X As String

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_AddItem(lstErr, cmdContext, "> SWAP : MT360-362-364  : export mail ...."): DoEvents
    Select Case cmdSelect_SQL_K
        Case "1":
            Call MSFlexGrid_SendMail(mMail_Destinataires, "SWAP", X, X, fgSelect, 7)
    End Select

Call lstErr_AddItem(lstErr, cmdContext, "Rétrocession de commissions  : export mail terminé"): DoEvents


Me.Enabled = True: Me.MousePointer = 0

End Sub







Public Sub ZSWIFTA0_Display()
Dim V, X As String
Dim xSql As String
On Error GoTo Error_Handler

currentAction = "ZSWIFTA0_Display"

xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIFTAF where SWIFTAETA = 1 and SWIFTANUM = " & mSWAMONZSWI
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
    If rsSab("SWIFTACOM") = "O" Then
        V = "message COMPLET" & vbCrLf & vbCrLf
    Else
        V = "message INCOMPLET" & vbCrLf & vbCrLf
    End If
    If rsSab("SWIFTASUP") = "O" Then
        V = V & "message SUPPRIME" & vbCrLf & vbCrLf
    End If
    If rsSab("SWIFTAVAL") = "O" Then
        V = V & "message VALIDE par " & rsSab("SWIFTAUT2") & vbCrLf & vbCrLf
    Else
        V = V & "message NON VALIDE" & vbCrLf & vbCrLf
    End If
    MsgBox V, vbInformation, "Message en attente SAB"
Else

    xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIHIAF where SWIHIAETA = 1 and SWIHIANUM = " & mSWAMONZSWI
    Set rsSab = cnsab.Execute(xSql)
    If Not rsSab.EOF Then
        If rsSab("SWIHIACOM") = "O" Then
            V = "message COMPLET" & vbCrLf & vbCrLf
        Else
            V = "message INCOMPLET" & vbCrLf & vbCrLf
        End If
        If rsSab("SWIHIASUP") = "O" Then
            V = V & "message SUPPRIME" & vbCrLf & vbCrLf
        End If
        If rsSab("SWIHIAVAL") = "O" Then
            V = V & "message VALIDE par " & rsSab("SWIHIAUT2") & vbCrLf & vbCrLf
        Else
            V = V & "message NON VALIDE" & vbCrLf & vbCrLf
        End If
        If rsSab("SWIHIADEN") > 0 Then
            V = V & "message ENVOYE le " & dateImp10_S(rsSab("SWIHIADEN") + 19000000) & " " & timeImp8(rsSab("SWIHIAHEN")) & vbCrLf
        Else
            V = V & "message NON ENVOYE" & vbCrLf
        End If
        MsgBox V, vbInformation, "Message Historisé SAB"
    End If
End If


Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fraSWI360_Display(lCLI As Long)
Dim V, X As String
Dim xSql As String
On Error GoTo Error_Handler

currentAction = "fraSWI360_Display"
blnSWI360_77H = False
blnSWI360_14C = False
blnSWI360_22B = False
X = Format(lCLI, "0000000")
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0," & paramIBM_Library_SAB & ".ZCLIENA0" _
     & " where BIATABID = 'SWI360' and BIATABK1 = '" & X & "' and CLIENACLI = BIATABK1 order by  BIATABK1 , BIATABK2"
Set rsSab = cnsab.Execute(xSql)


Do While Not rsSab.EOF
    Select Case Trim(rsSab("BIATABK2"))
        Case "77H": blnSWI360_77H = True: Call rsYBIATAB0_GetBuffer(rsSab, Old_YBIATAB0_77H)
                    txtSWI360_CLI = rsSab("CLIENACLI")
                    libSWI360_CLI = Trim(rsSab("CLIENARA1"))
                    txtSWI360_77H = Trim(rsSab("BIATABTXT"))
        Case "14C": blnSWI360_14C = True: Call rsYBIATAB0_GetBuffer(rsSab, Old_YBIATAB0_14C)
                    txtSWI360_14C = Trim(rsSab("BIATABTXT"))
        Case "22B": blnSWI360_22B = True: Call rsYBIATAB0_GetBuffer(rsSab, old_YBIATAB0_22B)
                    txtSWI360_22B = Trim(rsSab("BIATABTXT"))
    End Select
    
    rsSab.MoveNext

Loop


Set rsSab = Nothing
If blnSWI360_77H Or lCLI = 0 Then
    fraSWI360.Visible = True
Else
    V = "Manque l'enregistrement 77H"
    GoTo Error_MsgBox
End If
cmdSWI360_Update.Visible = blnSWI360_77H
cmdSWI360_Delete.Visible = blnSWI360_77H
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub txtSWI360_14C_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtSWI360_22B_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSWI360_77H_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSWI360_CLI_Change()
libSWI360_CLI = ""
End Sub

Private Sub txtSWI360_CLI_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub txtSWI360_CLI_LostFocus()
Dim xSql As String
xSql = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0" _
     & " where CLIENACLI = '" & Format(Val(txtSWI360_CLI), "0000000") & "'"
Set rsSab = cnsab.Execute(xSql)
If rsSab.EOF Then
    V = "Racine inconnue : " & txtSWI360_CLI
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Else
    libSWI360_CLI = Trim(rsSab("CLIENARA1"))
End If

End Sub


