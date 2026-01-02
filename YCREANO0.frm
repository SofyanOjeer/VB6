VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmYCREANO0 
   AutoRedraw      =   -1  'True
   Caption         =   "JPL"
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
   Icon            =   "YCREANO0.frx":0000
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
      TabPicture(0)   =   "YCREANO0.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "YCREANO0.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFg"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "YCREANO0.frx":0342
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
         TabIndex        =   18
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   22
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
         Text            =   "YCREANO0.frx":035E
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
         Begin VB.CommandButton cmdSAB_Dossier_DB 
            BackColor       =   &H0080C0FF&
            Caption         =   "Compta"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   9525
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1305
            Visible         =   0   'False
            Width           =   1125
         End
         Begin RichTextLib.RichTextBox txtRTF 
            Height          =   3735
            Left            =   960
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   7185
            Visible         =   0   'False
            Width           =   14775
            _ExtentX        =   26061
            _ExtentY        =   6588
            _Version        =   393217
            BackColor       =   14745599
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   3
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"YCREANO0.frx":0366
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
            Begin VB.ComboBox cboSelect_CREANOSTAK 
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
            Begin MSComCtl2.DTPicker txtSelect_CREANODCRE_Min 
               Height          =   300
               Left            =   4740
               TabIndex        =   14
               Top             =   480
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
               Format          =   100007939
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_CREANODCRE_Max 
               Height          =   300
               Left            =   6630
               TabIndex        =   15
               Top             =   495
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
               Format          =   100007939
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_CREANODCRE 
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
               Left            =   3750
               TabIndex        =   13
               Top             =   480
               Width           =   855
            End
            Begin VB.Label lblSelect_CREANOSTAK 
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
            Cols            =   12
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
            FormatString    =   $"YCREANO0.frx":03E6
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
      Picture         =   "YCREANO0.frx":0486
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
   Begin VB.Menu mnuCRE 
      Caption         =   "mnuCRE"
      Visible         =   0   'False
      Begin VB.Menu mnuCRE_LTXT 
         Caption         =   "Ajouter un commentaire"
      End
      Begin VB.Menu mnuCRE_Ann 
         Caption         =   "Annuler le suivi du CRE"
      End
      Begin VB.Menu mnuCRE_Restauration 
         Caption         =   "Restaurer le suivi du CRE"
      End
   End
End
Attribute VB_Name = "frmYCREANO0"
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

Dim xYCREANO0 As typeYCREANO0, oldYCREANO0 As typeYCREANO0, newYCREANO0 As typeYCREANO0
Dim mYCREANO0 As typeYCREANO0
Dim mYCREANO0_Update As String

Dim xYCRETXT0 As typeYCRETXT0, oldYCRETXT0 As typeYCRETXT0, newYCRETXT0 As typeYCRETXT0
Dim mYCRETXT0_Update As String
Dim HeightOfLine As Long, LinesOfText As Long

Dim txtRTF_prtForeColor_Header As Long


Dim VB_RTF_Modèle As String
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
    Call rsYCREANO0_GetBuffer(rsSab, xYCREANO0)
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

fgDetail.Col = 0: fgDetail.Text = xYCREANO0.CREANOSTAK
fgDetail.Col = 1: fgDetail.Text = xYCREANO0.CREANOSER & " " & xYCREANO0.CREANOSSE & " " & xYCREANO0.CREANOOPE & " " & xYCREANO0.CREANONUM & " " & xYCREANO0.CREANOEVE
fgDetail.Col = 2: fgDetail.Text = " " & dateImp10_S(xYCREANO0.CREANODCRE)
fgDetail.CellForeColor = vbMagenta
fgDetail.Col = 3: fgDetail.Text = Format(xYCREANO0.CREANOCRE, "### ### ###")
If xYCREANO0.CREANOSPLF > 0 Then fgDetail.Col = 4: fgDetail.Text = "oui"
If xYCREANO0.CREANOLTXT > 0 Then fgDetail.Col = 5: fgDetail.Text = "oui"
fgDetail.Col = 6: fgDetail.Text = xYCREANO0.CREANONB
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
Call DTPicker_Set(txtSelect_CREANODCRE_Min, DSys)
Call DTPicker_Set(txtSelect_CREANODCRE_Max, DSys)
txtSelect_CREANODCRE_Min.Value = Null

fgDetail_FormatString = fgDetail.FormatString
fgDetail.Enabled = True
fgDetail.Visible = False
fgDetail.Top = fgSelect.Top
fgDetail.Left = fgSelect.Left + fgSelect.Width - fgDetail.Width - 200

fraSelect_Options.Visible = True


If cboSelect_SQL.ListCount > 0 Then cboSelect_SQL.ListIndex = 0
cboSelect_CREANOSTAK.Clear
cboSelect_CREANOSTAK.AddItem "? CRE en anomalie"
cboSelect_CREANOSTAK.AddItem "A CRE annulés"
cboSelect_CREANOSTAK.AddItem "  CRE régularisés"
cboSelect_CREANOSTAK.AddItem "* tous les CRE"
cboSelect_CREANOSTAK.ListIndex = 0
blnControl = True

    '
txtRTF.LoadFile paramServer("\\BiaDoc\Filigrane\VB_RTF_Modèle.rtf")
VB_RTF_Modèle = txtRTF.TextRTF

txtRTF.Visible = False

Set fraUpdate.Container = fraSelect
fraUpdate.Top = fgSelect.Top
fraUpdate.Left = fgSelect.Left + fgSelect.Width - fraUpdate.Width - 200

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



Private Sub fgSelect_Display()

Dim K As Long

On Error GoTo Error_Handler
currentAction = "fgSelect_Display"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = fgSelect_FormatString

fgSelect.Rows = 1
                 
fgSelect.Row = 0

Do While Not rsSab.EOF

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    Call rsYCREANO0_GetBuffer(rsSab, xYCREANO0)
    fgSelect_Display_Line
    
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

Public Sub fgSelect_Display_Line()
Dim K As Integer, wColor As Long

On Error Resume Next

Select Case xYCREANO0.CREANOSTAK
    Case " ":
    Case "?": wColor = vbMagenta
    Case "A": wColor = RGB(160, 160, 160)
End Select
For K = 0 To 9
    fgSelect.Col = K: fgSelect.CellForeColor = wColor
    If xYCREANO0.CREANOSTAK = "?" Then fgSelect.CellFontBold = True

Next K
fgSelect.Col = 0: fgSelect.Text = xYCREANO0.CREANOSTAK
fgSelect.Col = 1: fgSelect.Text = xYCREANO0.CREANOSER & " " & xYCREANO0.CREANOSSE & " " & xYCREANO0.CREANOOPE & " " & xYCREANO0.CREANONUM & " " & xYCREANO0.CREANOEVE
fgSelect.Col = 2: fgSelect.Text = " " & dateImp10_S(xYCREANO0.CREANODCRE)
fgSelect.Col = 3: fgSelect.Text = Format(xYCREANO0.CREANOCRE, "### ### ###")
If xYCREANO0.CREANOSPLF > 0 Then
    fgSelect.Col = 4: fgSelect.Text = " oui"
    fgSelect.CellBackColor = mColor_W0
End If
If xYCREANO0.CREANOLTXT > 0 Then
    fgSelect.Col = 5: fgSelect.Text = "   oui"
    fgSelect.CellBackColor = mColor_W0
End If

fgSelect.Col = 6: fgSelect.Text = xYCREANO0.CREANONB
 If xYCREANO0.CREANODTRT > 0 Then fgSelect.Col = 7: fgSelect.Text = " " & dateImp10_S(xYCREANO0.CREANODTRT)
 If xYCREANO0.CREANOCREC > 0 Then fgSelect.Col = 8: fgSelect.Text = Format(xYCREANO0.CREANOCREC, "### ### ###")
 If xYCREANO0.CREANOPIE > 0 Then fgSelect.Col = 9: fgSelect.Text = Format(xYCREANO0.CREANOPIE, "### ### ###")
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
        Case "@CRE_ANO":    blnAuto = True
                        cmdSelect_SQL_K = "SPLF"
                        cmdSelect_Ok_Click
                        
                        mMail_Destinataires = srvSendMail.Exchange_Distribution("CRE_ANO", "@CRE_ANO")
                        cboSelect_CREANOSTAK.ListIndex = 0
                        txtSelect_CREANODCRE_Min.Value = Null
                        cmdSelect_SQL_K = "1"
                        cmdSelect_Ok_Click
                        mnuPrint_Mail_Click
                        Unload Me
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


Private Sub cboSelect_CREANOSTAK_Click()
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
Call frmSAB_Dossier_DB.Form_Init("MOUVEMDTR", "", 0, 0, mYCREANO0.CREANOSER, mYCREANO0.CREANOSSE, mYCREANO0.CREANOOPE, mYCREANO0.CREANONUM)
cmdSAB_Dossier_DB.Visible = False
End Sub

Private Sub cmdUpdate_Ok_Click()
Dim xSQL As String, xSet As String, X As String
On Error GoTo Error_Handler

currentAction = "màj CRE en anomalie"
Me.Enabled = False: Me.MousePointer = vbHourglass

newYCRETXT0.CRETXTINFO = Trim(txtUpdate_CRETXTINFO)
If newYCRETXT0.CRETXTINFO = "" Then
    Call MsgBox("Préciser un commentaire", vbExclamation, currentAction)
    GoTo Exit_sub
End If
newYCRETXT0.CRETXTYAMJ = DSys
newYCRETXT0.CRETXTYHMS = time_Hms
newYCRETXT0.CRETXTYUSR = usrName_UCase

Select Case mYCRETXT0_Update
    Case "New":  V = sqlYCRETXT0_Insert(newYCRETXT0): xSet = "CREANOLTXT = 1"
    Case "Update":  V = sqlYCRETXT0_Update(newYCRETXT0, oldYCRETXT0)
    Case Else:  V = "??? " & mYCRETXT0_Update
End Select
If Not IsNull(V) Then GoTo Error_MsgBox

If mYCREANO0_Update = "Update" Then
    Select Case mYCREANO0.CREANOSTAK
        Case "?": X = "A"
        Case "A": X = "?"
    End Select
    If xSet = "" Then
        xSet = "CREANOSTAK = '" & X & "'"
    Else
        xSet = xSet & " , CREANOSTAK = '" & X & "'"
    End If
End If
If xSet <> "" Then
    xSQL = "Update " & paramIBM_Library_SABSPE & ".YCREANO0 " _
         & " set " & xSet & " where CREANOETA = 1  and CREANOCRE = " & mYCREANO0.CREANOCRE
    Set rsSab = cnsab.Execute(xSQL)
End If
fraUpdate.Visible = False
Call cmdSelect_SQL_1
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
                fgDetail.Col = 1: wX = Trim(fgDetail.Text)
                K = 0
                mYCREANO0.CREANOSER = Space_Scan(wX, K)
                mYCREANO0.CREANOSSE = Space_Scan(wX, K)
                mYCREANO0.CREANOOPE = Space_Scan(wX, K)
                mYCREANO0.CREANONUM = Space_Scan(wX, K)
                mYCREANO0.CREANOEVE = Space_Scan(wX, K)
                fgDetail.Col = 3: mYCREANO0.CREANOCRE = Val(fgDetail.Text)
                fgDetail.Col = 4: mYCREANO0.CREANOSPLF = IIf(Trim(fgDetail.Text) = "", 0, 1)
                fgDetail.Col = 5: mYCREANO0.CREANOLTXT = IIf(Trim(fgDetail.Text) = "", 0, 1)
                fgDetail.Col = 9: mYCREANO0.CREANOPIE = Val(fgDetail.Text)
                fraDetail_Display_txtRTF
                
            Case "2"
                fgDetail.Col = 1: wX = Trim(fgDetail.Text)
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
                fgSelect.Col = 1: wX = Trim(fgSelect.Text)
                K = 0
                mYCREANO0.CREANOSER = Space_Scan(wX, K)
                mYCREANO0.CREANOSSE = Space_Scan(wX, K)
                mYCREANO0.CREANOOPE = Space_Scan(wX, K)
                mYCREANO0.CREANONUM = Space_Scan(wX, K)
                mYCREANO0.CREANOEVE = Space_Scan(wX, K)
                fgSelect.Col = 3: mYCREANO0.CREANOCRE = Val(fgSelect.Text)
                fgSelect.Col = 4: mYCREANO0.CREANOSPLF = IIf(Trim(fgSelect.Text) = "", 0, 1)
                fgSelect.Col = 5: mYCREANO0.CREANOLTXT = IIf(Trim(fgSelect.Text) = "", 0, 1)
                fgSelect.Col = 9: mYCREANO0.CREANOPIE = Val(fgSelect.Text)
                fraDetail_Display
                If arrHab(2) Then
                    mnuCRE_Ann.Visible = False
                    mnuCRE_Restauration.Visible = False
                    fgSelect.Col = 0: mYCREANO0.CREANOSTAK = Trim(fgSelect.Text)
                    If mYCREANO0.CREANOSTAK = "?" Then mnuCRE_Ann.Visible = True
                    If mYCREANO0.CREANOSTAK = "A" Then mnuCRE_Restauration.Visible = True
                    Me.PopupMenu mnuCRE, vbPopupMenuLeftButton
                End If
            Case "2"
                fgSelect.Col = 1: wX = Trim(fgSelect.Text)
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
    
    Select Case cmdSelect_SQL_K
        Case "1": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
        Case "SPLF": cmdSelect_Ok.Visible = True
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
cmdSelect_Clear

If Mid$(cboSelect_CREANOSTAK, 1, 1) = "*" Then
    If Not IsNull(txtSelect_CREANODCRE_Min.Value) Then
        Call DTPicker_Control(txtSelect_CREANODCRE_Min, wAmjMin)
        Call DTPicker_Control(txtSelect_CREANODCRE_Max, wAmjMax)
        xWhere = "where CREANODCRE >= " & wAmjMin & " And CREANODCRE <= " & wAmjMax
    End If
Else
    xWhere = "where CREANOSTAK = '" & Mid$(cboSelect_CREANOSTAK, 1, 1) & "'"
    If Not IsNull(txtSelect_CREANODCRE_Min.Value) Then
        Call DTPicker_Control(txtSelect_CREANODCRE_Min, wAmjMin)
        Call DTPicker_Control(txtSelect_CREANODCRE_Max, wAmjMax)
        xWhere = xWhere & " and CREANODCRE >= " & wAmjMin & " And CREANODCRE <= " & wAmjMax
    End If
End If

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCREANOLK " & xWhere & " order by CREANODCRE , CREANOCRE"

Set rsSab = cnsab.Execute(xSQL)
  

fgSelect_Display

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
Call lstErr_Clear(lstErr, cmdContext, "> CRE_ANO_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Clear

Select Case cmdSelect_SQL_K
    Case "1": cmdSelect_SQL_1
    Case "SPLF": cmdSelect_SQL_SPLF
    Case "JPL": cmdSelect_SQL_JPL '
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< CRE_ANO_cmdSelect_Ok"): DoEvents
lstErr.Height = 480
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
cmdSelect_Ok.BackColor = fgSelect.BackColorFixed
End Sub



Private Sub mnuCRE_Ann_Click()
mYCREANO0_Update = "Update"
cmdUpdate_Ok.Caption = "ANNULER le suivi"
cmdUpdate_Ok.BackColor = mColor_W1
fraUpdate_Display
End Sub

Private Sub mnuCRE_LTXT_Click()
mYCREANO0_Update = ""
cmdUpdate_Ok.Caption = "Enregistrer"
cmdUpdate_Ok.BackColor = mColor_G2
fraUpdate_Display

End Sub


Private Sub mnuCRE_Restauration_Click()
mYCREANO0_Update = "Update"
cmdUpdate_Ok.Caption = "Restaurer"
cmdUpdate_Ok.BackColor = mColor_W1
fraUpdate_Display

End Sub

Private Sub mnuPrint_Excel_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String
Call lstErr_AddItem(lstErr, cmdContext, "> CRE_ANO : export Excel ...."): DoEvents
    Select Case cmdSelect_SQL_K
        Case "1":
            X = "Situation des CRE en anomalie au " & dateImp10_S(DSys) & " " & Time
            Call MSflexGrid_Excel("", "CRE_ANO", X, fgSelect, 9)
    End Select

Call lstErr_AddItem(lstErr, cmdContext, "< CRE_ANO : export Excel terminé"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint_Mail_Click()
Dim X As String

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_AddItem(lstErr, cmdContext, "> CRE_ANO : export mail ...."): DoEvents
    Select Case cmdSelect_SQL_K
        Case "1":
            X = "Situation des CRE en anomalie au " & dateImp10_S(DSys) & " " & Time
            Call MSFlexGrid_SendMail(mMail_Destinataires, "CRE_ANO", X, X, fgSelect, 9)
    End Select

Call lstErr_AddItem(lstErr, cmdContext, "< CRE_ANO : export mail terminé"): DoEvents


Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub txtSelect_CREANODCRE_Max_Change()
cmdSelect_Clear

End Sub


Private Sub txtSelect_CREANODCRE_Max_Click()
cmdSelect_Clear

End Sub


Private Sub txtSelect_CREANODCRE_Min_Change()
cmdSelect_Clear

End Sub


Private Sub txtSelect_CREANODCRE_Min_Click()
cmdSelect_Clear

End Sub



Public Sub cmdSelect_SQL_SPLF()
Dim objFolder, objFiles, fsoFile As File
Dim mCREANOCRE_Orphelin As Long, intFile As Integer, xIn As String, X As String, K As Integer
Dim blnDeleteFile As Boolean, xSQL As String
On Error GoTo Error_Handler
'
Me.Enabled = False
Screen.MousePointer = vbHourglass

Call rsYCREANO0_Init(xYCREANO0)
xYCREANO0.CREANOETA = 1
xYCREANO0.CREANOAGE = 1
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCREANO0 " _
     & " where CREANOETA = 1 and CREANOCRE < 5000 order by CREANOCRE desc"

Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    mCREANOCRE_Orphelin = rsSab("CREANOCRE")
Else
    mCREANOCRE_Orphelin = 0
End If

Set objFolder = msFileSystem.GetFolder(paramZSCHCRO0_SPLF)
Set objFiles = objFolder.Files
For Each fsoFile In objFiles
    blnDeleteFile = False
    xYCREANO0.CREANONUM = 0
    intFile = FreeFile(0)
    Open paramZSCHCRO0_SPLF & fsoFile.Name For Input As #intFile
    Do Until EOF(1)
        DoEvents
        Line Input #intFile, xIn
        If InStr(xIn, "SAB073U01") > 0 Then
            blnDeleteFile = True
            Exit Do
        End If
        K = InStr(xIn, "SAB07301     SIEGE")
        If K > 0 Then
            K = K + 18
            X = Space_Scan(xIn, K) & Space(10)
            xYCREANO0.CREANODCRE = Mid$(X, 7, 4) & Mid$(X, 4, 2) & Mid$(X, 1, 2)
            K = InStr(xIn, "SCHGE005P1/")
            If K > 0 Then
                K = K + 11
                xYCREANO0.CREANOSER = Mid$(xIn, K, 2)
                xYCREANO0.CREANOSSE = Mid$(xIn, K + 2, 2)
        End If

        End If
        K = InStr(xIn, "Références schéma")
        If K > 0 Then
            K = InStr(xIn, "Opération")
            If K > 0 Then
                K = K + 9
                xYCREANO0.CREANOOPE = Space_Scan(xIn, K)
                xYCREANO0.CREANONUM = Val(Space_Scan(xIn, K))
                K = InStr(K, xIn, "Evénement")
                If K > 0 Then
                    K = K + 9
                    xYCREANO0.CREANOEVE = Space_Scan(xIn, K)
                    Exit Do
                End If
            End If
        End If
    Loop
    Close intFile
    If blnDeleteFile Then msFileSystem.DeleteFile fsoFile.Path, True
    If xYCREANO0.CREANONUM > 0 Then
    
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCREANOLA " _
             & " where CREANOETA = 1 and CREANOAGE = 1" _
             & " and CREANOSER = '" & xYCREANO0.CREANOSER & "' and CREANOSSE = '" & xYCREANO0.CREANOSSE & "'" _
             & " and CREANOOPE = '" & xYCREANO0.CREANOOPE & "' and CREANONUM = " & xYCREANO0.CREANONUM _
             & " and CREANOEVE = '" & xYCREANO0.CREANOEVE & "' and CREANOSPLF = 0" _
             & " and CREANODCRE = " & xYCREANO0.CREANODCRE

        Set rsSab = cnsab.Execute(xSQL)
        If Not rsSab.EOF Then
            xYCREANO0.CREANOCRE = rsSab("CREANOCRE")
            xSQL = "Update " & paramIBM_Library_SABSPE & ".YCREANO0 " _
                 & " set CREANOSPLF = 1 where CREANOETA = 1  and CREANOCRE = " & xYCREANO0.CREANOCRE
            
            Set rsSab = cnsab.Execute(xSQL)
                 
            msFileSystem.MoveFile fsoFile.Path, paramYCREANO0 & "CRE_" & xYCREANO0.CREANOCRE & ".txt"
        Else
        
            xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCREANOH " _
                 & " where  CREANODCRE = " & xYCREANO0.CREANODCRE _
                 & " and CREANOOPE = '" & xYCREANO0.CREANOOPE & "' and CREANONUM = " & xYCREANO0.CREANONUM _
                & " and CREANOEVE = '" & xYCREANO0.CREANOEVE & "' and CREANOSPLF = 0"
            Set rsSab = cnsab.Execute(xSQL)
            If Not rsSab.EOF Then
                xYCREANO0.CREANOCRE = rsSab("CREANOCRE")
                xSQL = "Update " & paramIBM_Library_SABSPE & ".YCREANOH " _
                     & " set CREANOSPLF = 1 where CREANOETA = 1  and CREANOCRE = " & xYCREANO0.CREANOCRE
                
                Set rsSab = cnsab.Execute(xSQL)
                     
                msFileSystem.MoveFile fsoFile.Path, paramYCREANO0 & "CRE_" & xYCREANO0.CREANOCRE & ".txt"
                
            Else
                xYCREANO0.CREANOSTAK = "?"
                xYCREANO0.CREANOSPLF = 1
                mCREANOCRE_Orphelin = mCREANOCRE_Orphelin + 1
                xYCREANO0.CREANOCRE = mCREANOCRE_Orphelin
                V = sqlYCREANO0_Insert(xYCREANO0)
                If IsNull(V) Then
                    msFileSystem.MoveFile fsoFile.Path, paramYCREANO0 & "CRE_" & xYCREANO0.CREANOCRE & ".txt"
                End If
            End If
        End If
    End If
        

Next
Screen.MousePointer = vbDefault
Me.Enabled = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdSelect_SQL_SPLF " & fsoFile.Path


End Sub

Public Sub fraDetail_Display()
Dim xSQL As String

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCREANOH " _
     & " where CREANOETA = 1 and CREANOAGE = 1 " _
     & " and CREANOSER ='" & mYCREANO0.CREANOSER & "' and CREANOSSE='" & mYCREANO0.CREANOSSE & "'" _
     & " and CREANOOPE ='" & mYCREANO0.CREANOOPE & "' and CREANONUM=" & mYCREANO0.CREANONUM _
     & " and CREANOEVE ='" & mYCREANO0.CREANOEVE & "'" _
    & " order by CREANODCRE , CREANOCRE"
Set rsSab = cnsab.Execute(xSQL)
Call fgDetail_Display
If fgDetail.Rows = 1 Then fgDetail.Visible = False

cmdSAB_Dossier_DB.Visible = IIf(mYCREANO0.CREANOPIE > 0, True, False)
fraDetail_Display_txtRTF
End Sub
Public Sub fraDetail_Display_txtRTF()
Dim xRTF As String, intFile As Integer, xIn As String

If mYCREANO0.CREANOLTXT > 0 Then
    Dim xSQL As String
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCRETXT0 " _
         & " where CRETXTETA = 1 " _
         & " and CRETXTCRE =" & mYCREANO0.CREANOCRE
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        Call rsYCRETXT0_GetBuffer(rsSab, xYCRETXT0)
         xRTF = "\fs16\highlight12\cf13 Commentaire de : \cf10 " & xYCRETXT0.CRETXTYUSR _
         & " le " & dateImp10_S(xYCRETXT0.CRETXTYAMJ) & " " & timeImp8(xYCRETXT0.CRETXTYHMS) _
         & "  \highlight9\cf9 _____________________________________________________________________________________\par" _
         & "\highlight0\fs18\par\tab\cf13 " _
         & Replace(xYCRETXT0.CRETXTINFO, vbCrLf, "\par\tab\cf13 ") & "\highlight9 "
    End If
End If
If mYCREANO0.CREANOSPLF > 0 Then
    intFile = FreeFile(0)
    xRTF = xRTF & "\fs16\cf1\par"
    Open paramYCREANO0 & "CRE_" & mYCREANO0.CREANOCRE & ".txt" For Input As #intFile
    Do Until EOF(1)
        DoEvents
        
        Line Input #intFile, xIn
        If Mid$(xIn, 1, 1) <> "$" Then
            Mid$(xIn, 1, 4) = "    "
            Select Case Mid$(xIn, 6, 3)
                Case "BAS": xIn = "\cf14 " & xIn
                Case "Réf": xIn = "\par\cf13 " & xIn & "\par"
                Case Else: xIn = "\cf1 " & xIn
            End Select
            If Mid$(xIn, 10, 10) = "SCHGE005P1" Then
                xIn = Replace(xIn, "SIEGE        ", "SIEGE        \highlight12\cf10 ")
                xIn = Replace(xIn, "PAGE", "\highlight0\cf1 PAGE ")
            End If
            xRTF = xRTF & "\par " & xIn
        End If
    Loop
    Close intFile
End If
If xRTF <> "" Then
   txtRTF.TextRTF = VB_RTF_Modèle
   txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", xRTF)
   Call txtRTF_Visible
End If
End Sub


Public Sub txtRTF_Visible()
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "{\f0\fswiss\fprq2\fcharset0 Calibri;}", "{\f0\fmodern\fprq1\fcharset0 Courier New;}")
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "\cf1\f0\fs20 a\cf2 b\cf3 c\cf4 d\cf5 e\cf6 f\cf7 g\cf8 h\cf9 i\cf10 j\cf11 k\cf12 l\cf13 m\cf14 n\cf15 o\cf16 p", "")
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", "")
txtRTF.Visible = True

End Sub


Public Sub fraUpdate_Display()
Dim xSQL As String

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCRETXT0 " _
     & " where CRETXTETA = 1 " _
     & " and CRETXTCRE =" & mYCREANO0.CREANOCRE
Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    mYCRETXT0_Update = "New"
    newYCRETXT0.CRETXTETA = 1
    newYCRETXT0.CRETXTCRE = mYCREANO0.CREANOCRE
    newYCRETXT0.CRETXTINFO = ""
    txtUpdate_CRETXTINFO = ""
Else
    mYCRETXT0_Update = "Update"
    Call rsYCRETXT0_GetBuffer(rsSab, oldYCRETXT0)
    txtUpdate_CRETXTINFO = oldYCRETXT0.CRETXTINFO
    newYCRETXT0 = oldYCRETXT0
End If


fraUpdate.Visible = True
End Sub

Public Sub cmdSelect_SQL_JPL()
Dim xSQL As String

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCREANOLK " _
     & " where  CREANOETA = 1 and CREANOSTAK ='?' order by CREANODCRE" _

Set rsSab_X = cnsab.Execute(xSQL)
Do While Not rsSab_X.EOF

    Call rsYCREANO0_GetBuffer(rsSab_X, mYCREANO0)
    If mYCREANO0.CREANODCRE < 20130000 Then
        newYCRETXT0.CRETXTETA = mYCREANO0.CREANOETA
        newYCRETXT0.CRETXTCRE = mYCREANO0.CREANOCRE
        mYCRETXT0_Update = "New": txtUpdate_CRETXTINFO = "Reprise : annulation automatique"
        mYCREANO0_Update = "Update"
        Call cmdUpdate_Ok_Click
    End If
    
    rsSab_X.MoveNext

Loop

End Sub
