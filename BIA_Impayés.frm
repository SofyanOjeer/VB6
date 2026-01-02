VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBIA_Impayés 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA_Impayés"
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
   Icon            =   "BIA_Impayés.frx":0000
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
      Left            =   8925
      TabIndex        =   2
      Top             =   45
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   11640
      Left            =   15
      TabIndex        =   3
      Top             =   480
      Width           =   16275
      _ExtentX        =   28707
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
      TabPicture(0)   =   "BIA_Impayés.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "BIA_Impayés.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraSelect_Options_1der"
      Tab(1).Control(1)=   "txtFg"
      Tab(1).Control(2)=   "txtRTF"
      Tab(1).Control(3)=   "fgDetail"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "BIA_Impayés.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame fraSelect_Options_1der 
         BackColor       =   &H00F0FFFF&
         Caption         =   "Liste des clients débiteurs"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   -74220
         TabIndex        =   12
         Top             =   8295
         Visible         =   0   'False
         Width           =   12075
         Begin VB.TextBox txtSelect_Racine 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   13
            Top             =   555
            Width           =   1305
         End
         Begin MSComCtl2.DTPicker txtSelect_AMJMIN 
            Height          =   300
            Left            =   6480
            TabIndex        =   16
            Top             =   555
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
            Format          =   105054211
            CurrentDate     =   41302
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblSelect_AMJMIN 
            BackColor       =   &H00F0FFFF&
            Caption         =   "date de début de recherche du débit"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3150
            TabIndex        =   15
            Top             =   585
            Width           =   3120
         End
         Begin VB.Label lblSelect_Racine 
            BackColor       =   &H00F0FFFF&
            Caption         =   "racine"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   14
            Top             =   570
            Width           =   825
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
         Text            =   "BIA_Impayés.frx":035E
         Top             =   1155
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Frame fraSelect 
         BackColor       =   &H00E0E0E0&
         Height          =   11055
         Left            =   -105
         TabIndex        =   4
         Top             =   495
         Width           =   16290
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
            Left            =   14640
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   810
            Width           =   1335
         End
         Begin VB.ComboBox cboSelect_SQL 
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
            Left            =   12315
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   285
            Width           =   3840
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
            Height          =   1305
            Left            =   165
            TabIndex        =   5
            Top             =   90
            Visible         =   0   'False
            Width           =   12105
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   9585
            Left            =   135
            TabIndex        =   10
            Top             =   1410
            Width           =   16095
            _ExtentX        =   28390
            _ExtentY        =   16907
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
            FormatString    =   $"BIA_Impayés.frx":0366
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin RichTextLib.RichTextBox txtRTF 
         Height          =   5610
         Left            =   -67590
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1740
         Visible         =   0   'False
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   9895
         _Version        =   393217
         BackColor       =   15790320
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"BIA_Impayés.frx":0463
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
         Height          =   7680
         Left            =   -74355
         TabIndex        =   11
         Top             =   510
         Visible         =   0   'False
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   13547
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         RowHeightMin    =   400
         BackColor       =   16448250
         ForeColor       =   4210752
         BackColorFixed  =   12640511
         ForeColorFixed  =   0
         BackColorBkg    =   16448250
         GridColor       =   10526720
         GridColorFixed  =   10526720
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         FormatString    =   $"BIA_Impayés.frx":04E3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Left            =   15795
      Picture         =   "BIA_Impayés.frx":05B0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   15
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
Attribute VB_Name = "frmBIA_Impayés"
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

Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long



Dim HeightOfLine As Long, LinesOfText As Long

Dim txtRTF_prtForeColor_Header As Long

Dim rsSab_2 As New ADODB.Recordset, rsSab_3 As New ADODB.Recordset
Dim mSQL_PLANCOPRO As String
Dim meCV1 As typeCV, meCV2 As typeCV

Dim mSOLDECEN_EUR As Currency, mAUTSYCMON_EUR As Currency, blnZAUTSYC0 As Boolean, blnZAUTSYC0_Existe As Boolean
Dim mSOLDECEN_Nbj As Integer
Dim xYBIACPT0 As typeYBIACPT0
Dim xZAUTSYC0 As typeZAUTSYC0
Dim mAMJ_Min As Long, mDate_Min As String, mIBM_Min As Long
Dim mAMJ_Max As Long, mDate_Max As String, mIBM_Max As Long
Dim mAMJ_3M As Long, mDate_3M As String, m3M_Nbj As Integer

Dim mIBM_MOUVEMDVA As Long, m365_NbJ As Long, m365_Lib As String

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

Private Sub fgDetail_Display_1(lCLIENACLI As String)
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
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where CLIENACLI = '" & lCLIENACLI & "' and " & mSQL_PLANCOPRO

Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    Call rsYBIACPT0_GetBuffer(rsSab, xYBIACPT0)
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_Display_1_Line
    rsSab.MoveNext
Loop
'DR 30/01/2013
rsSab.Close
Set rsSab = Nothing
'FIN DR 30/01/2013
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



Public Sub fgDetail_Display_1_Line()
Dim X As String, wAmj As Long

On Error Resume Next
fgDetail.Col = 0: fgDetail.Text = Trim(xYBIACPT0.COMPTECOM)
fgDetail.Col = 1: fgDetail.Text = Trim(xYBIACPT0.COMPTEINT)
fgDetail.Col = 2: fgDetail.Text = xYBIACPT0.PLANCOPRO
meCV1.Montant = -xYBIACPT0.SOLDECEN

fgDetail.Col = 3: fgDetail.Text = Format(meCV1.Montant, "### ### ### ##0.00")
If meCV1.Montant < 0 Then fgDetail.CellForeColor = vbRed
fgDetail.Col = 4: fgDetail.Text = xYBIACPT0.COMPTEDEV
fgDetail.Col = 4: fgDetail.CellFontBold = True: fgDetail.CellForeColor = vbBlue
If meCV1.Montant <> 0 Then
    meCV1.DeviseIso = xYBIACPT0.COMPTEDEV
    If meCV1.DeviseIso <> "EUR" Then
        meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
        Call CV_Calc("", meCV1, meCV2)
    Else
        meCV2.Montant = meCV1.Montant
    End If
    fgDetail.Col = 5: fgDetail.Text = Format(meCV2.Montant, "### ### ### ##0.00")
    If meCV2.Montant < 0 Then fgDetail.CellForeColor = vbRed
End If
If xYBIACPT0.SOLDEDMO <> 0 Then
    fgDetail.Col = 6: fgDetail.Text = dateImp10(xYBIACPT0.SOLDEDMO + 19000000)
End If
fgDetail.Col = 7: fgDetail.Text = xYBIACPT0.COMPTEFON

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
fgDetail.Left = 3500

fraSelect_Options.Visible = True

fraSelect_Options_1der.Visible = False
Set fraSelect_Options_1der.Container = fraSelect
fraSelect_Options_1der.Top = fraSelect_Options.Top
fraSelect_Options_1der.Left = fraSelect_Options.Left
fraSelect_Options_1der.Height = fraSelect_Options.Height
fraSelect_Options_1der.Width = fraSelect_Options.Width

fgDetail.Visible = False
Set fgDetail.Container = fraSelect
fgDetail.Top = fgSelect.Top
fgDetail.Height = fgSelect.Height
fgDetail.Left = fgSelect.Left + fgSelect.Width - fgDetail.Width + 100




mAMJ_Max = YBIATAB0_DATE_CPT_J: mDate_Max = Date_VB(mAMJ_Max, 0): mIBM_Max = mAMJ_Max - 19000000
mDate_Min = DateAdd("d", -365, mDate_Max)
mAMJ_Min = Mid$(mDate_Min, 7, 4) & Mid$(mDate_Min, 4, 2) & Mid$(mDate_Min, 1, 2)
mIBM_Min = mAMJ_Min - 19000000
mDate_3M = DateAdd("m", -3, mDate_Max)
mAMJ_3M = Mid$(mDate_3M, 7, 4) & Mid$(mDate_3M, 4, 2) & Mid$(mDate_3M, 1, 2)
m3M_Nbj = DateDiff("d", mDate_3M, mDate_Max)

X = DateAdd("d", -3650, mDate_Max)
X = Mid$(X, 7, 4) & Mid$(X, 4, 2) & Mid$(X, 1, 2)

Call DTPicker_Set(txtSelect_AMJMIN, X)
txtSelect_AMJMIN.Enabled = False

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

Dim K As Long, xSQL As String, blnRupture As Boolean, blnDisplay As Boolean
Dim numdossier As String
Dim dateimpaye As String

On Error GoTo Error_Handler
currentAction = "fgSelect_Display"
fgSelect.Visible = True ' False
fgSelect_Reset
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Rows = 1
                 
fgSelect.Row = 0
fgSelect.Col = 0: fgSelect.CellAlignment = 1 'flexAlignLeftCenter
fgSelect.Col = 1: fgSelect.CellAlignment = 1 'flexAlignLeftCenter
fgSelect.Col = 2: fgSelect.CellAlignment = 1 'flexAlignLeftCenter
fgSelect.Col = 3: fgSelect.CellAlignment = 7 'flexAlignRightCenter
fgSelect.ColAlignment(4) = 4
fgSelect.Col = 4: fgSelect.CellAlignment = 4 'flexAlignCenterCenter
fgSelect.ColAlignment(5) = 4
fgSelect.Col = 5: fgSelect.CellAlignment = 4 'flexAlignCenterCenter
fgSelect.Col = 6: fgSelect.CellAlignment = 7 'flexAlignRightCenter
fgSelect.ColAlignment(7) = 4
fgSelect.Col = 7: fgSelect.CellAlignment = 4 'flexAlignCenterCenter
If cmdSelect_SQL_K = "2" Then
    fgSelect.Col = 6
    fgSelect.Text = "Montant du crédit"
    fgSelect.Col = 7
    fgSelect.Text = "Fin du crédit"
    fgSelect.ColAlignment(8) = 1
    fgSelect.Col = 8: fgSelect.CellAlignment = 1 'flexAlignLeftCenter
    fgSelect.Text = "Dossier"
    fgSelect.ColWidth(8) = fgSelect.ColWidth(4)
    numdossier = ""
    dateimpaye = ""
End If

Do While Not rsSab.EOF
    mSOLDECEN_EUR = 0
    blnRupture = False
    blnDisplay = False
    blnZAUTSYC0 = False: blnZAUTSYC0_Existe = False
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
         & " where CLIENACLI = '" & rsSab("CLIENACLI") & "' and " & mSQL_PLANCOPRO

    Set rsSab_2 = cnsab.Execute(xSQL)
    Do While Not rsSab_2.EOF
        If Not blnRupture Then blnRupture = True: Call rsYBIACPT0_GetBuffer(rsSab_2, xYBIACPT0)
        If fctUser_Classe_Aut(xYBIACPT0.COMPTECLA) Then
            meCV1.Montant = -rsSab_2("SOLDECEN") / 1000
        Else
            meCV1.Montant = 999999999999.99
        End If
        If meCV1.Montant <> 0 Then
            meCV1.DeviseIso = rsSab_2("COMPTEDEV")
            If meCV1.DeviseIso <> "EUR" Then
                meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
                Call CV_Calc("", meCV1, meCV2)
            Else
                meCV2.Montant = meCV1.Montant
            End If
            mSOLDECEN_EUR = mSOLDECEN_EUR + meCV2.Montant
        End If
        rsSab_2.MoveNext
    Loop
    'DR 30/01/2013
    rsSab_2.Close
    Set rsSab_2 = Nothing
    'FIN DR 30/01/2013
    If cmdSelect_SQL_K = "2" Then
        'trouver le numéro du dossier pour ce client
        xSQL = "select IMPECHNUM, IMPECHDTI from " & paramIBM_Library_SAB & ".ZIMPECH0 where impechpay='" & rsSab("CLIENACLI") & "' and impechcoe=2"
        Set rsSab_2 = cnsab.Execute(xSQL)
        Do While Not rsSab_2.EOF
            numdossier = rsSab_2("IMPECHNUM")
            dateimpaye = rsSab_2("IMPECHDTI")
            rsSab_2.MoveNext
        Loop
        rsSab_2.Close
        Set rsSab_2 = Nothing
    End If
'______________________________________________________________________________________________________
    If mSOLDECEN_EUR < 0 Then
        blnDisplay = True
        If cmdSelect_SQL_K <> "2" Then
            X = "select * from " & paramIBM_Library_SAB & ".ZAUTSYC0 " _
                & " where AUTSYCTYP = '1' and AUTSYCAUT = 'DEC' and AUTSYCCLI = '" & rsSab("CLIENACLI") & "'"
            Set rsSab_2 = cnsab.Execute(X)
            If Not rsSab_2.EOF Then
                 blnZAUTSYC0_Existe = True
               Call rsZAUTSYC0_GetBuffer(rsSab_2, xZAUTSYC0)
                mAUTSYCMON_EUR = xZAUTSYC0.AUTSYCMON
                If mAUTSYCMON_EUR = 0 Then
                   ' blnZAUTSYC0 = False   '$JPL 2015-01-23
                Else
                    blnZAUTSYC0 = True
                    If xZAUTSYC0.AUTSYCDEV <> "EUR" Then
                       meCV1.Montant = xZAUTSYC0.AUTSYCMON
                       meCV1.DeviseIso = xZAUTSYC0.AUTSYCDEV
                       meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
                       Call CV_Calc("", meCV1, meCV2)
                       mAUTSYCMON_EUR = meCV2.Montant
                    End If
                End If
                If xZAUTSYC0.AUTSYCFIN + 19000000 < YBIATAB0_DATE_CPT_J Then
                    blnDisplay = True
                Else
                    If mAUTSYCMON_EUR + mSOLDECEN_EUR < 0 Then
                        blnDisplay = True
                    Else
                        blnDisplay = False
                    End If
                End If
            End If
            'DR 30/01/2013
            rsSab_2.Close
            Set rsSab_2 = Nothing
            'FIN DR 30/01/2013
        End If
        
        If cmdSelect_SQL_K = "1der" Or cmdSelect_SQL_K = "2" Then blnDisplay = True
        
        If blnDisplay < 0 Then
            Call fgSelect_Display_1_MOUVEMDVA
            'Call fgSelect_Display_1_MOUVEMDTR ' contrôle par date de traitement qui reflète la situation réelle
            'contrairement à date comptable à cause des écritures complémentaires
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            Call fgSelect_Display_1_Line(numdossier, dateimpaye)
        End If
    End If
    rsSab.MoveNext

Loop

fgSelect.Visible = True


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSelect_Display_1_Line(numdossier As String, dateimpaye As String)
Dim X As String, wAmj As Long, blnA_Déclasser As Boolean, wColor As Long

On Error Resume Next
If numdossier <> "" And dateimpaye <> "" Then
    fgSelect.Col = 8: fgSelect.Text = numdossier
End If
blnA_Déclasser = False
fgSelect.Col = 0: fgSelect.Text = xYBIACPT0.CLIENACLI
fgSelect.Col = 1: fgSelect.Text = xYBIACPT0.CLIENARES
fgSelect.Col = 2: fgSelect.Text = Trim(xYBIACPT0.CLIENARA1) & " " & Trim(xYBIACPT0.CLIENARA2)
fgSelect.Col = 3: fgSelect.Text = Format(mSOLDECEN_EUR, "### ### ### ##0.00")
If mSOLDECEN_EUR < 0 Then fgSelect.CellForeColor = vbRed
fgSelect.Col = 4:
If mSOLDECEN_Nbj = m365_NbJ Then
    blnA_Déclasser = True
    fgSelect.Text = m365_Lib
    fgSelect.CellBackColor = mColor_Y2
Else
    fgSelect.Text = mSOLDECEN_Nbj
    Select Case mSOLDECEN_Nbj
        Case Is < 32: fgSelect.CellBackColor = mColor_Y0: wColor = mColor_Y0
        Case Is < 63: fgSelect.CellBackColor = mColor_Y1: wColor = mColor_Y1
        Case Is < m3M_Nbj: fgSelect.CellBackColor = mColor_Y2: wColor = mColor_Y2
        Case Else: blnA_Déclasser = True: fgSelect.CellBackColor = mColor_Y3: wColor = mColor_Y3
    End Select
    fgSelect.Col = 5: fgSelect.Text = "   " & DateAdd("d", -mSOLDECEN_Nbj, mDate_Max)
End If

If blnA_Déclasser Then
    
    If Abs(mSOLDECEN_EUR) > 200 Then
        wColor = mColor_W0
        Dim K As Integer
        For K = 0 To 5: fgSelect.Col = K: fgSelect.CellBackColor = mColor_W0: fgSelect.CellFontBold = True: Next K
    Else
        fgSelect.Col = 3: fgSelect.CellBackColor = wColor
    End If
    
End If
fgSelect.Col = 5: fgSelect.CellBackColor = wColor
'If blnZAUTSYC0 Then
If blnZAUTSYC0_Existe Then
    If blnZAUTSYC0 Then
        fgSelect.Col = 6: fgSelect.Text = Format(mAUTSYCMON_EUR, "### ### ### ##0.00")
        fgSelect.CellBackColor = wColor
        wAmj = xZAUTSYC0.AUTSYCFIN + 19000000
        fgSelect.Col = 7: fgSelect.Text = "  " & dateImp10(wAmj)
        If wAmj < YBIATAB0_DATE_CPT_J Then fgSelect.CellForeColor = vbMagenta
        fgSelect.CellBackColor = wColor
    Else
        fgSelect.Col = 7
        If xZAUTSYC0.AUTSYCFIN > 0 Then
            fgSelect.Text = "  " & dateImp10(xZAUTSYC0.AUTSYCFIN + 19000000)
        Else
            fgSelect.Text = "  X"
        End If
        fgSelect.CellBackColor = RGB(228, 228, 228)
    End If
End If
If cmdSelect_SQL_K = "2" And numdossier <> "" Then
    'trouver la date de l'échéance impayée
    'trouver montant du crédit et date de fin du crédit
    X = "SELECT CREDOSMNT, CREDOSDFI, CREDOSDEV from " & paramIBM_Library_SAB & ".ZCREDOS0 where CREDOSDOS='" & numdossier & "'"
    Set rsSab_2 = cnsab.Execute(X)
    Do While Not rsSab_2.EOF
        fgSelect.Col = 5
        fgSelect.Text = Date_VB(CDbl(dateimpaye) + 19000000, 0)
        fgSelect.Col = 4
        fgSelect.Text = CDbl(DateValue(Now)) - CDbl(DateValue(fgSelect.Text))
        fgSelect.Col = 6
        fgSelect.Text = Format(CDbl(rsSab_2("CREDOSMNT")), "### ### ### ##0.00") & " " & rsSab_2("CREDOSDEV")
        fgSelect.Col = 7
        fgSelect.Text = Date_VB(CDbl(rsSab_2("CREDOSDFI")) + 19000000, 0)
        rsSab_2.MoveNext
    Loop
    rsSab_2.Close
    Set rsSab_2 = Nothing
End If
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim wFct As String

mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

wFct = UCase$(Trim(Mid$(Msg, 1, 12)))
Call BIA_VB_HAB(wFct, arrHab(), cboSelect_SQL)

Call Form_Init
Select Case wFct
    Case "@BIA_IMPAYÉS"
        Dim xDest As String
             blnAuto = True
            cmdSelect_SQL_K = "1"
            cmdSelect_Ok_Click
            X = "Etat des clients débiteurs en compte courant au " & mDate_Max
            xDest = srvSendMail.Exchange_Distribution("BIA_IMPAYÉS", "@BIA_IMPAYÉS")
            Call MSFlexGrid_SendMail(xDest, "@BIA_IMPAYÉS", X, X, fgSelect, fgSelect.Cols - 1)
            
            'désactivé le 21/11/2018 à la demande R. Benmalek
            'réactivé le 15/01/2019 à la demande R. Benmalek
            cmdSelect_SQL_K = "2"
            cmdSelect_Ok_Click
            X = "Etat des impayés sur crédit au " & mDate_Max
            Call MSFlexGrid_SendMail(xDest, "@BIA_IMPAYÉS", X, X, fgSelect, fgSelect.Cols - 1)
            
            cmdSelect_SQL_K = "1dcom"
            cmdSelect_Ok_Click
            X = "Etat des clients débiteurs en compte courant (R00-R59) au " & mDate_Max
            xDest = srvSendMail.Exchange_Distribution("BIA_IMPAYÉS", "@DCOM")
            Call MSFlexGrid_SendMail(xDest, "@BIA_IMPAYÉS", X, X, fgSelect, fgSelect.Cols - 1)
            
'            'désactivé le 21/11/2018 à la demande R. Benmalek
'            cmdSelect_SQL_K = "2"
'            cmdSelect_Ok_Click
'            X = "Etat des impayés sur crédit au " & mDate_Max
'            xDest = "rosillette.d@bia-paris.fr"
'            Call MSFlexGrid_SendMail(xDest, "@BIA_IMPAYÉS", X, X, fgSelect, fgSelect.Cols - 1)
            
            cmdSelect_SQL_K = "1jur"
            cmdSelect_Ok_Click
            X = "Etat des clients débiteurs en compte courant (R80-R89) au " & mDate_Max
            xDest = srvSendMail.Exchange_Distribution("BIA_IMPAYÉS", "@JURIDIQUE")
            Call MSFlexGrid_SendMail(xDest, "@BIA_IMPAYÉS", X, X, fgSelect, fgSelect.Cols - 1)
            
            cmdSelect_SQL_K = "RIAD"
            cmdSelect_Ok_Click
            X = "Etat des clients débiteurs en compte courant (classe de sécurité 60, 61, 62) au " & mDate_Max
            xDest = srvSendMail.Exchange_Distribution("BIA_IMPAYÉS", "@PERSONNEL")
            Call MSFlexGrid_SendMail(xDest, "@BIA_IMPAYÉS", X, X, fgSelect, fgSelect.Cols - 1)
            
            Unload Me
    Case Else: blnAuto = False: cboSelect_SQL.ListIndex = 0

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

Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String
On Error Resume Next


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
        fgDetail.Col = 0: wX = Trim(fgDetail.Text)
        Call frmSAB_Dossier_DB.Form_Init("MOUVEMDVA", wX, CStr(mIBM_MOUVEMDVA + 19000000), CStr(mAMJ_Max), "", "", "", 0)
   End If
End If
fgDetail.LeftCol = 0


End Sub




Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, xUUMID As String
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
            Case "1", "1dcom", "1jur", "1der"
                fgSelect.Col = 0: wX = Trim(fgSelect.Text)
                Call fgDetail_Display_1(wX)
            Case "2"
                fgSelect.Col = 0: wX = Trim(fgSelect.Text)
                Call fgDetail_Display_1(wX)
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
cmdSelect_Ok.BackColor = vbGreen

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
    fraSelect_Options_1der.Visible = False
    
    Select Case cmdSelect_SQL_K
        Case "1", "1dcom", "1jur": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
        Case "2": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
        Case "1der": cmdSelect_Ok.Visible = True: fraSelect_Options_1der.Visible = True
    End Select

End If
End Sub


Private Sub cmdSelect_SQL_2()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1"
xWhere = ""

Set rsSab = cnsab.Execute(xSQL)
  

'DR 30/01/2013
rsSab.Close
'FIN DR 30/01/2013


Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_1()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1"

mIBM_MOUVEMDVA = mIBM_Min
m365_NbJ = 365: m365_Lib = "> 1an"

xSQL = "select distinct CLIENARES , CLIENACLI from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where " & mSQL_PLANCOPRO _
     & " and COMPTECOM in (select SOLDECOM  from " & paramIBM_Library_SAB & ".ZSOLDE0  where SOLDEETA = 1 and SOLDEPLA = 1  and SOLDECEN > 0 )" _
     & " order by CLIENARES , CLIENACLI"

Set rsSab = cnsab.Execute(xSQL)
  

fgSelect_Display_1

'DR 30/01/2013
rsSab.Close
'FIN DR 30/01/2013

Set rsSab = Nothing

Exit Sub

Error_Handler:
    
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_1der()
Dim V, X As String, wAmj As Long
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1der"

Call DTPicker_Control(txtSelect_AMJMIN, X)
'wAMJ = Mid$(X, 7, 4) & Mid$(X, 4, 2) & Mid$(X, 1, 2)

mIBM_MOUVEMDVA = X - 19000000
m365_NbJ = 3650: m365_Lib = "> 10 ans"


If Trim(txtSelect_Racine) = "" Then
    X = ""
Else
    X = "and CLIENACLI =  '" & Format(txtSelect_Racine, "0000000") & "'"
End If

xSQL = "select distinct CLIENARES , CLIENACLI from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where " & mSQL_PLANCOPRO & X _
     & " and COMPTECOM in (select SOLDECOM  from " & paramIBM_Library_SAB & ".ZSOLDE0  where SOLDEETA = 1 and SOLDEPLA = 1  and SOLDECEN > 0 )" _
     & " order by CLIENARES , CLIENACLI"

Set rsSab = cnsab.Execute(xSQL)
  

fgSelect_Display_1

'DR 30/01/2013
rsSab.Close
'FIN DR 30/01/2013

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
Call lstErr_Clear(lstErr, cmdContext, "> BIA_Impayés_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Clear

Select Case cmdSelect_SQL_K
    Case "1":
            mSQL_PLANCOPRO = " PLANCOPRO in ( 'CAV' , 'DOR', 'LOR' , 'LOB' ) and CLIENARES <> 'R60' AND COMPTECLA not in ('60','61','62')"
            cmdSelect_SQL_1
    Case "1dcom":
            mSQL_PLANCOPRO = " PLANCOPRO in ( 'CAV' , 'DOR', 'LOR' , 'LOB' ) and CLIENARES between 'R00' and 'R59' AND COMPTECLA not in ('60','61','62')"
            cmdSelect_SQL_1
     Case "1jur":
            mSQL_PLANCOPRO = " PLANCOPRO in ( 'CAV' , 'DOR', 'LOR' , 'LOB' ) and CLIENARES between 'R80' and 'R89' AND COMPTECLA not in ('60','61','62')"
            cmdSelect_SQL_1
   Case "2":
            mSQL_PLANCOPRO = " PLANCOPRO in ( 'IMP' ) AND COMPTECLA not in ('60','61','62')"
            cmdSelect_SQL_1
    Case "1der":
            mSQL_PLANCOPRO = " PLANCOPRO in ( 'CAV' , 'DOR', 'LOR' , 'LOB' ) and CLIENARES <> 'R60' AND COMPTECLA not in ('60','61','62')"
            cmdSelect_SQL_1der
    Case "RIAD":
            mSQL_PLANCOPRO = " PLANCOPRO in ( 'CAV' , 'DOR', 'LOR' , 'LOB' ) AND COMPTECLA in ('60','61','62')"
            cmdSelect_SQL_1
    Case "JPL":
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_Impayés_cmdSelect_Ok"): DoEvents
lstErr.Height = 480
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
cmdSelect_Ok.BackColor = fgSelect.BackColorFixed
End Sub




Public Sub fgSelect_Display_1_MOUVEMDVA()

'$JPL 2015-01-22 correction bug : ajout arrAUT()

Dim K As Long
Dim arrSolde(3660) As Currency, arrAut(3660) As Currency
Dim wAmj As Long, wDate As String, wIBM As Long

X = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where CLIENACLI = '" & xYBIACPT0.CLIENACLI & "' and " & mSQL_PLANCOPRO

Set rsSab_2 = cnsab.Execute(X)
Do While Not rsSab_2.EOF

    X = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
         & " where MOUVEMCOM = '" & rsSab_2("COMPTECOM") & "' and MOUVEMDVA >= " & mIBM_MOUVEMDVA _
         & " order by MOUVEMDVA"
    
    Set rsSab_3 = cnsab.Execute(X)
    Do While Not rsSab_3.EOF
        If wIBM <> rsSab_3("MOUVEMDVA") Then
            wIBM = rsSab_3("MOUVEMDVA")
            wAmj = wIBM + 19000000
            meCV1.OpéAmj = wAmj
            K = DateDiff("d", Date_VB(wAmj, 0), mDate_Max)
            If K < 0 Then K = 0
        End If
        curX = -rsSab_3("MOUVEMMON")
        'calcul dans la devise d'origine
        If rsSab_3("COMPTEDEV") <> "EUR" Then
           meCV1.Montant = curX
           meCV1.DeviseIso = rsSab_3("COMPTEDEV")
           Call CV_Calc("", meCV1, meCV2)
           curX = meCV2.Montant
           If meCV1.Cours = 0 And wAmj > YBIATAB0_DATE_CPT_J Then  ' cours  non connu à J
                meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
                Call CV_Calc("", meCV1, meCV2)
                curX = meCV2.Montant
           End If
        End If
        arrSolde(K) = arrSolde(K) + curX
        rsSab_3.MoveNext
    Loop
    'DR 30/01/2013
    rsSab_3.Close
    Set rsSab_3 = Nothing
    'FIN DR 30/01/2013
    
    
    rsSab_2.MoveNext
Loop
'DR 30/01/2013
rsSab_2.Close
Set rsSab_2 = Nothing
'FIN DR 30/01/2013

If cmdSelect_SQL_K = "1der" Then
Else

    If blnZAUTSYC0 Then
        Dim K1 As Long, K2 As Long
        wAmj = xZAUTSYC0.AUTSYCDEB + 19000000
        K2 = DateDiff("d", Date_VB(wAmj, 0), mDate_Max)
        If K2 > m365_NbJ Then K2 = m365_NbJ
        wAmj = xZAUTSYC0.AUTSYCFIN + 19000000
        K1 = DateDiff("d", Date_VB(wAmj, 0), mDate_Max)
        If K1 < 0 Then K1 = 0
        For K = K1 To K2
            arrAut(K) = -mAUTSYCMON_EUR   '$JPL 2013-01-23
           'arrSolde(K) = arrSolde(K) - mAUTSYCMON_EUR
        Next K
'$JPL 2013-01-23 historique des autorisations
        X = "select * from " & paramIBM_Library_SAB & ".ZAUTHST0 " _
            & " where AUTHSTTYP = '1' and AUTHSTAUT = 'DEC' and AUTHSTCLI = '" & rsSab("CLIENACLI") & "'" _
            & " and AUTHSTFIN > " & mIBM_Min & " order by AUTHSTFIN"
        
        Set rsSab_2 = cnsab.Execute(X)
        
        Do While Not rsSab_2.EOF
        
            wAmj = rsSab_2("AUTHSTDEB") + 19000000
            K2 = DateDiff("d", Date_VB(wAmj, 0), mDate_Max)
            If K2 > m365_NbJ Then K2 = m365_NbJ
            wAmj = rsSab_2("AUTHSTFIN") + 19000000
            K1 = DateDiff("d", Date_VB(wAmj, 0), mDate_Max)
            If K1 < 0 Then K1 = 0
            For K = K1 To K2
                arrAut(K) = -rsSab_2("AUTHSTMON")
            Next K
            
            rsSab_2.MoveNext
        
            Loop

    End If
End If

mSOLDECEN_Nbj = m365_NbJ
curX = mSOLDECEN_EUR
For K = 0 To m365_NbJ
    curX = curX - arrSolde(K)
    'If curX >= 0 Then
    If curX >= arrAut(K) Then
        mSOLDECEN_Nbj = K + 1
        Exit For
    End If
Next K
End Sub

Public Sub fgSelect_Display_1_MOUVEMDCO()

'$JPL 2015-01-22 correction bug : ajout arrAUT()

Dim K As Long
Dim arrSolde(3660) As Currency, arrAut(3660) As Currency
Dim wAmj As Long, wDate As String, wIBM As Long

X = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where CLIENACLI = '" & xYBIACPT0.CLIENACLI & "' and " & mSQL_PLANCOPRO

Set rsSab_2 = cnsab.Execute(X)
Do While Not rsSab_2.EOF

    X = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
         & " where MOUVEMCOM = '" & rsSab_2("COMPTECOM") & "' and MOUVEMDCO >= " & mIBM_MOUVEMDVA _
         & " order by MOUVEMDCO"
    
    Set rsSab_3 = cnsab.Execute(X)
    Do While Not rsSab_3.EOF
        If wIBM <> rsSab_3("MOUVEMDCO") Then
            wIBM = rsSab_3("MOUVEMDCO")
            wAmj = wIBM + 19000000
            meCV1.OpéAmj = wAmj
            K = DateDiff("d", Date_VB(wAmj, 0), mDate_Max)
            If K < 0 Then K = 0
        End If
        curX = -rsSab_3("MOUVEMMON")
        'calcul dans la devise d'origine
        If rsSab_3("COMPTEDEV") <> "EUR" Then
           meCV1.Montant = curX
           meCV1.DeviseIso = rsSab_3("COMPTEDEV")
           Call CV_Calc("", meCV1, meCV2)
           curX = meCV2.Montant
           If meCV1.Cours = 0 And wAmj > YBIATAB0_DATE_CPT_J Then  ' cours  non connu à J
                meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
                Call CV_Calc("", meCV1, meCV2)
                curX = meCV2.Montant
           End If
        End If
        arrSolde(K) = arrSolde(K) + curX
        rsSab_3.MoveNext
    Loop
    'DR 30/01/2013
    rsSab_3.Close
    Set rsSab_3 = Nothing
    'FIN DR 30/01/2013
    
    
    rsSab_2.MoveNext
Loop
'DR 30/01/2013
rsSab_2.Close
Set rsSab_2 = Nothing
'FIN DR 30/01/2013

If cmdSelect_SQL_K = "1der" Then
Else

    If blnZAUTSYC0 Then
        Dim K1 As Long, K2 As Long
        wAmj = xZAUTSYC0.AUTSYCDEB + 19000000
        K2 = DateDiff("d", Date_VB(wAmj, 0), mDate_Max)
        If K2 > m365_NbJ Then K2 = m365_NbJ
        wAmj = xZAUTSYC0.AUTSYCFIN + 19000000
        K1 = DateDiff("d", Date_VB(wAmj, 0), mDate_Max)
        If K1 < 0 Then K1 = 0
        For K = K1 To K2
            arrAut(K) = -mAUTSYCMON_EUR   '$JPL 2013-01-23
           'arrSolde(K) = arrSolde(K) - mAUTSYCMON_EUR
        Next K
'$JPL 2013-01-23 historique des autorisations
        X = "select * from " & paramIBM_Library_SAB & ".ZAUTHST0 " _
            & " where AUTHSTTYP = '1' and AUTHSTAUT = 'DEC' and AUTHSTCLI = '" & rsSab("CLIENACLI") & "'" _
            & " and AUTHSTFIN > " & mIBM_Min & " order by AUTHSTFIN"
        
        Set rsSab_2 = cnsab.Execute(X)
        
        Do While Not rsSab_2.EOF
        
            wAmj = rsSab_2("AUTHSTDEB") + 19000000
            K2 = DateDiff("d", Date_VB(wAmj, 0), mDate_Max)
            If K2 > m365_NbJ Then K2 = m365_NbJ
            wAmj = rsSab_2("AUTHSTFIN") + 19000000
            K1 = DateDiff("d", Date_VB(wAmj, 0), mDate_Max)
            If K1 < 0 Then K1 = 0
            For K = K1 To K2
                arrAut(K) = -rsSab_2("AUTHSTMON")
            Next K
            
            rsSab_2.MoveNext
        
            Loop

    End If
End If

mSOLDECEN_Nbj = m365_NbJ
curX = mSOLDECEN_EUR
For K = 0 To m365_NbJ
    curX = curX - arrSolde(K)
    'If curX >= 0 Then
    If curX >= arrAut(K) Then
        mSOLDECEN_Nbj = K + 1
        Exit For
    End If
Next K
End Sub
Public Sub fgSelect_Display_1_MOUVEMDTR()

'$JPL 2015-01-22 correction bug : ajout arrAUT()

Dim K As Long
Dim arrSolde(3660) As Currency, arrAut(3660) As Currency
Dim wAmj As Long, wDate As String, wIBM As Long

X = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where CLIENACLI = '" & xYBIACPT0.CLIENACLI & "' and " & mSQL_PLANCOPRO

Set rsSab_2 = cnsab.Execute(X)
Do While Not rsSab_2.EOF

    X = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
         & " where MOUVEMCOM = '" & rsSab_2("COMPTECOM") & "' and MOUVEMDTR >= " & mIBM_MOUVEMDVA _
         & " order by MOUVEMDTR"
    
    Set rsSab_3 = cnsab.Execute(X)
    Do While Not rsSab_3.EOF
        If wIBM <> rsSab_3("MOUVEMDTR") Then
            wIBM = rsSab_3("MOUVEMDTR")
            wAmj = wIBM + 19000000
            meCV1.OpéAmj = wAmj
            K = DateDiff("d", Date_VB(wAmj, 0), mDate_Max)
            If K < 0 Then K = 0
        End If
        curX = -rsSab_3("MOUVEMMON")
        'calcul dans la devise d'origine
        If rsSab_3("COMPTEDEV") <> "EUR" Then
           meCV1.Montant = curX
           meCV1.DeviseIso = rsSab_3("COMPTEDEV")
           Call CV_Calc("", meCV1, meCV2)
           curX = meCV2.Montant
           If meCV1.Cours = 0 And wAmj > YBIATAB0_DATE_CPT_J Then  ' cours  non connu à J
                meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
                Call CV_Calc("", meCV1, meCV2)
                curX = meCV2.Montant
           End If
        End If
        arrSolde(K) = arrSolde(K) + curX
        rsSab_3.MoveNext
    Loop
    'DR 30/01/2013
    rsSab_3.Close
    Set rsSab_3 = Nothing
    'FIN DR 30/01/2013
    
    
    rsSab_2.MoveNext
Loop
'DR 30/01/2013
rsSab_2.Close
Set rsSab_2 = Nothing
'FIN DR 30/01/2013

If cmdSelect_SQL_K = "1der" Then
Else

    If blnZAUTSYC0 Then
        Dim K1 As Long, K2 As Long
        wAmj = xZAUTSYC0.AUTSYCDEB + 19000000
        K2 = DateDiff("d", Date_VB(wAmj, 0), mDate_Max)
        If K2 > m365_NbJ Then K2 = m365_NbJ
        wAmj = xZAUTSYC0.AUTSYCFIN + 19000000
        K1 = DateDiff("d", Date_VB(wAmj, 0), mDate_Max)
        If K1 < 0 Then K1 = 0
        For K = K1 To K2
            arrAut(K) = -mAUTSYCMON_EUR   '$JPL 2013-01-23
           'arrSolde(K) = arrSolde(K) - mAUTSYCMON_EUR
        Next K
'$JPL 2013-01-23 historique des autorisations
        X = "select * from " & paramIBM_Library_SAB & ".ZAUTHST0 " _
            & " where AUTHSTTYP = '1' and AUTHSTAUT = 'DEC' and AUTHSTCLI = '" & rsSab("CLIENACLI") & "'" _
            & " and AUTHSTFIN > " & mIBM_Min & " order by AUTHSTFIN"
        
        Set rsSab_2 = cnsab.Execute(X)
        
        Do While Not rsSab_2.EOF
        
            wAmj = rsSab_2("AUTHSTDEB") + 19000000
            K2 = DateDiff("d", Date_VB(wAmj, 0), mDate_Max)
            If K2 > m365_NbJ Then K2 = m365_NbJ
            wAmj = rsSab_2("AUTHSTFIN") + 19000000
            K1 = DateDiff("d", Date_VB(wAmj, 0), mDate_Max)
            If K1 < 0 Then K1 = 0
            For K = K1 To K2
                arrAut(K) = -rsSab_2("AUTHSTMON")
            Next K
            
            rsSab_2.MoveNext
        
            Loop

    End If
End If

mSOLDECEN_Nbj = m365_NbJ
curX = mSOLDECEN_EUR
For K = 0 To m365_NbJ
    curX = curX - arrSolde(K)
    'If curX >= 0 Then
    If curX >= arrAut(K) Then
        mSOLDECEN_Nbj = K + 1
        Exit For
    End If
Next K
End Sub

Private Sub mnuPrint_Excel_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String
Call lstErr_AddItem(lstErr, cmdContext, "> BIA_Impayés : export Excel ...."): DoEvents
Select Case cmdSelect_SQL_K
    Case "1", "1dcom", "1jur":
        X = cmdSelect_SQL_K & " - Etat des clients EN DEPASSEMENT en compte courant au " & mDate_Max
        Call MSflexGrid_Excel("", "BIA_Impayés", X, fgSelect, fgSelect.Cols - 1)
    Case "1der":
        X = cmdSelect_SQL_K & " - Etat des clients DEBITEURS en compte courant au " & mDate_Max
        Call MSflexGrid_Excel("", "BIA_Impayés", X, fgSelect, fgSelect.Cols - 1)
    Case "2":
        X = "Etat des impayés sur crédit au " & mDate_Max
        Call MSflexGrid_Excel("", "BIA_Impayés", X, fgSelect, fgSelect.Cols - 1)
End Select
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_Impayés : export Excel terminé"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuPrint_Mail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String
Call lstErr_AddItem(lstErr, cmdContext, "> BIA_Impayés : export mail ...."): DoEvents

Select Case cmdSelect_SQL_K
    Case "1", "1dcom", "1jur":
        X = cmdSelect_SQL_K & " - Etat des clients EN DEPASSEMENT en compte courant au " & mDate_Max
        Call MSFlexGrid_SendMail(currentSSIWINMAIL, "BIA_Impayés", X, X, fgSelect, fgSelect.Cols - 1)
    Case "1der":
        X = cmdSelect_SQL_K & " - Etat des clients DEBITEURS en compte courant au " & mDate_Max
        Call MSFlexGrid_SendMail(currentSSIWINMAIL, "BIA_Impayés", X, X, fgSelect, fgSelect.Cols - 1)
    Case "2":
        X = "Etat des impayés sur crédit au " & mDate_Max
        Call MSFlexGrid_SendMail(currentSSIWINMAIL, "BIA_Impayés", X, X, fgSelect, fgSelect.Cols - 1)
End Select

Call lstErr_AddItem(lstErr, cmdContext, "< BIA_Impayés : export mail terminé"): DoEvents


Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub txtSelect_Racine_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


