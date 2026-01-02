VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSAB_Stock 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_Stock"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13560
   Icon            =   "SAB_Stock.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9150
   ScaleWidth      =   13560
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   4
      Top             =   0
      Width           =   5175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   2
      Top             =   500
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   15266
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Rechercher"
      TabPicture(0)   =   "SAB_Stock.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Dossiers"
      TabPicture(1)   =   "SAB_Stock.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraYBIASTO0"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "'"
      TabPicture(2)   =   "SAB_Stock.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame fraYBIASTO0 
         Height          =   8025
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   13290
         Begin MSFlexGridLib.MSFlexGrid fgYBIASTO0 
            Height          =   7125
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   13200
            _ExtentX        =   23283
            _ExtentY        =   12568
            _Version        =   393216
            Rows            =   1
            Cols            =   9
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   14737632
            ForeColor       =   8388608
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"SAB_Stock.frx":035E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label libYBIASTO0_YSTOMON 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   2655
         End
         Begin VB.Label libYBIASTO0_SOLDECEN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label libYBIASTO0_Diff 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5640
            TabIndex        =   10
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label libYBIASTO0_Total 
            Caption         =   "Total"
            Height          =   255
            Left            =   3000
            TabIndex        =   9
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label libYBIASTO0_Solde 
            Caption         =   "Solde"
            Height          =   255
            Left            =   2880
            TabIndex        =   8
            Top             =   600
            Width           =   2415
         End
      End
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   13290
         Begin VB.Frame fraSelect_Options 
            Height          =   1005
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   11355
            Begin VB.CheckBox chkSelect_YSTOMON 
               Caption         =   "afficher les dossiers dont Encours =0"
               Height          =   315
               Left            =   8040
               TabIndex        =   28
               Top             =   480
               Width           =   3255
            End
            Begin VB.CheckBox chkSelect_Ecart 
               Caption         =   "afficher uniquement les comptes en écart "
               Height          =   315
               Left            =   8040
               TabIndex        =   27
               Top             =   120
               Value           =   1  'Checked
               Width           =   3255
            End
            Begin VB.ComboBox cboSelect_YSTOCLI 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   6480
               Sorted          =   -1  'True
               TabIndex        =   26
               Text            =   "CLI"
               Top             =   520
               Width           =   1300
            End
            Begin VB.ComboBox cboSelect_YSTONAT 
               Height          =   315
               Left            =   6480
               Sorted          =   -1  'True
               TabIndex        =   24
               Text            =   "NAT"
               Top             =   120
               Width           =   1300
            End
            Begin VB.ComboBox cboSelect_YSTOAPP 
               Height          =   315
               Left            =   1200
               Sorted          =   -1  'True
               TabIndex        =   21
               Text            =   "APP"
               Top             =   120
               Width           =   1300
            End
            Begin VB.ComboBox cboSelect_YSTOOPE 
               Height          =   315
               Left            =   3840
               Sorted          =   -1  'True
               TabIndex        =   17
               Text            =   "OPE"
               Top             =   120
               Width           =   1300
            End
            Begin VB.ComboBox cboSelect_YSTODEV 
               Height          =   315
               Left            =   1200
               Sorted          =   -1  'True
               TabIndex        =   16
               Text            =   "DEV"
               Top             =   520
               Width           =   1300
            End
            Begin VB.ComboBox cboSelect_YSTOPCI 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   3840
               Sorted          =   -1  'True
               TabIndex        =   15
               Text            =   "PCI"
               Top             =   520
               Width           =   1300
            End
            Begin VB.Label lblSelect_YSTOCLI 
               Caption         =   "Client"
               Height          =   210
               Left            =   5640
               TabIndex        =   25
               Top             =   560
               Width           =   540
            End
            Begin VB.Label lblSelect_YSTONAT 
               Caption         =   "Nature"
               Height          =   270
               Left            =   5400
               TabIndex        =   23
               Top             =   240
               Width           =   690
            End
            Begin VB.Label lblSelect_YSTOAPP 
               Caption         =   "Application"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   840
            End
            Begin VB.Label lblSelect_YSTOPCI 
               Caption         =   "PCI"
               Height          =   210
               Left            =   2760
               TabIndex        =   20
               Top             =   560
               Width           =   540
            End
            Begin VB.Label lblSelect_YSTOOPE 
               Caption         =   "Opération"
               Height          =   270
               Left            =   2760
               TabIndex        =   19
               Top             =   240
               Width           =   690
            End
            Begin VB.Label lblSelect_YSTODEV 
               Caption         =   "Devise"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   600
               Width           =   600
            End
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   645
            Left            =   11880
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6585
            Left            =   120
            TabIndex        =   5
            Top             =   1200
            Width           =   13080
            _ExtentX        =   23072
            _ExtentY        =   11615
            _Version        =   393216
            Rows            =   1
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   500
            BackColor       =   14737632
            ForeColor       =   4210688
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   "<Compte              |<Intitulé                               |> Dev  |>Solde / Encours   |>Ecart              |"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
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
      Picture         =   "SAB_Stock.frx":0418
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Print_Liste 
         Caption         =   "Imprimer liste"
      End
      Begin VB.Menu mnuSelect_Print_Détail 
         Caption         =   "Imprimer liste détaillée"
      End
   End
End
Attribute VB_Name = "frmSAB_Stock"
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
Dim SAB_Stock_Aut As typeAuthorization
Dim curX1 As Currency, curX2 As Currency

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim fgYBIASTO0_FormatString As String, fgYBIASTO0_K As Integer
Dim fgYBIASTO0_RowDisplay As Integer, fgYBIASTO0_RowClick As Integer, fgYBIASTO0_ColClick As Integer
Dim fgYBIASTO0_ColorClick As Long, fgYBIASTO0_ColorDisplay As Long
Dim fgYBIASTO0_Sort1 As Integer, fgYBIASTO0_Sort2 As Integer
Dim fgYBIASTO0_SortAD As Integer, fgYBIASTO0_Sort1_Old As Integer
Dim fgYBIASTO0_arrIndex As Integer
Dim blnfgYBIASTO0_DisplayLine As Boolean
Dim meYBIASTO0 As typeYBIASTO0, xYBIASTO0 As typeYBIASTO0
Dim arrYBIASTO0() As typeYBIASTO0
Dim selYBIASTO0() As typeYBIASTO0, selYBIASTO0_Nb As Long, selYBIASTO0_Max As Long


Dim meYBIACPT0 As typeYBIACPT0, xYBIACPT0 As typeYBIACPT0
Dim arrYBIACPT0() As typeYBIACPT0, arrCompte_Nb As Long, arrCompte_Max As Long

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel

Public Sub cmdPrint_Ok_xlsManual(lFct As String, wsExcel As Excel.Worksheet)
Dim nbRows As Long

fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Liste : " & fgSelect.Rows - 1)

nbRows = prtSAB_Stock_Monitor_xlsManual(Trim(lFct), fgSelect, arrYBIASTO0(), arrYBIACPT0(), arrCompte_Nb, wsExcel)
If nbRows > 0 Then
    Call zoneImpression_xlsManual(Trim(lFct), nbRows, wsExcel)
    Call wsExcel.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, paramIMP_PDF_Path & "\" & paramEditionNoPaper_Auto_PgmName & ".pdf", XlFixedFormatQuality.xlQualityMinimum, False)
    'sauvegarde du fichier
    Call impressions_xlsManual.prtIMP_PDF_Monitor_xlsManual
End If

fgSelect.Visible = True
Me.Show

End Sub
Private Sub fgYBIASTO0_Montant_Negatif()
'xxxx Modification montant négatif 21/12/2009 Denis R.
'Affiche les lignes négatives en rouge
Dim I As Long
Dim J As Long

    For I = 1 To fgYBIASTO0.Rows - 1
        fgYBIASTO0.Row = I: fgYBIASTO0.Col = 1
        If CDbl(fgYBIASTO0.Text) < 0 Then
            For J = 0 To fgYBIASTO0.Cols - 1
                fgYBIASTO0.Col = J
                fgYBIASTO0.CellForeColor = vbRed
            Next J
        End If
    Next I
    
End Sub

Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSelect.Row

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
        fgSelect.Col = 0
    End If
End If

End Sub
Private Sub fgSelect_Display()
Dim V
Dim X As String, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim I As Integer, K As Long

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
cmdPrint.Enabled = False

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgSelect_Display"

For I = 1 To arrCompte_Nb
    blnDisplay = True
    xYBIASTO0 = arrYBIASTO0(I)
    xYBIACPT0 = arrYBIACPT0(I)
    If Mid$(xYBIASTO0.YSTONAT, 1, 3) = "PAR" Then
        If Not SAB_Stock_Aut.Valider Then blnDisplay = False
    End If
    
    If blnDisplay Then
        If fctUser_Classe_Aut(xYBIACPT0.COMPTECLA) Then fgSelect_DisplayLine I
    End If
Next I


fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
fgSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Comptes : " & arrCompte_Nb): DoEvents
If fgSelect.Rows > 1 Then
    cmdPrint.Enabled = True
Else
    If chkSelect_Ecart.Value = "1" Then MsgBox "NEANT", vbInformation, "SAB_Stock : liste des écarts"
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub


Public Sub fgYBIASTO0_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgYBIASTO0.Row

If lRow > 0 And lRow < fgYBIASTO0.Rows Then
    fgYBIASTO0.Row = lRow
    For I = 0 To fgYBIASTO0_arrIndex
        fgYBIASTO0.Col = I: fgYBIASTO0.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgYBIASTO0.Row = mRow
    If fgYBIASTO0.Row > 0 Then
        lRow = fgYBIASTO0.Row
        lColor_Old = fgYBIASTO0.CellBackColor
        For I = 0 To fgYBIASTO0_arrIndex
          fgYBIASTO0.Col = I: fgYBIASTO0.CellBackColor = lColor
        Next I
        fgYBIASTO0.Col = 0
    End If
End If

End Sub

Private Sub fgYBIASTO0_Display()
Dim xSQL As String
Dim intReturn As Integer
Dim curTotal As Currency, nbTotal As Long, curSolde As Currency
fgYBIASTO0_Reset
fgYBIASTO0.Rows = 1
fgYBIASTO0.FormatString = fgYBIASTO0_FormatString

ReDim selYBIASTO0(101)
selYBIASTO0_Max = 99: selYBIASTO0_Nb = 0

Set rsSab = Nothing
libYBIASTO0_Diff = ""
curTotal = 0: nbTotal = 0
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIASTO0 where " _
     & "YSTOPCI = '" & xYBIASTO0.YSTOPCI & "'" _
     & "AND YSTODEV = '" & xYBIASTO0.YSTODEV & "'" _
     & "AND YSTOCLI = " & xYBIASTO0.YSTOCLI

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYBIASTO0_GetBuffer(rsSab, xYBIASTO0)
    If Not IsNull(V) Then
        MsgBox V, vbCritical, "frmSAB_Balance.SQL_ODBC"
        Exit Sub
    Else
        nbTotal = nbTotal + 1
        fgYBIASTO0_DisplayLine nbTotal
        curTotal = curTotal + xYBIASTO0.YSTOMON
        
        selYBIASTO0_Nb = selYBIASTO0_Nb + 1
        If selYBIASTO0_Nb > selYBIASTO0_Max Then
            selYBIASTO0_Max = selYBIASTO0_Max + 100
            ReDim Preserve selYBIASTO0(selYBIASTO0_Max)
        End If
        
       selYBIASTO0(selYBIASTO0_Nb) = xYBIASTO0

    End If
    rsSab.MoveNext
Loop
libYBIASTO0_Total = nbTotal & " dossiers : " & xYBIACPT0.COMPTEDEV
libYBIASTO0_Solde = "Solde " & xYBIACPT0.COMPTEDEV & " au " & dateIBM10(YBIATAB0_DATE_CPT_J, True)

libYBIASTO0_YSTOMON = Format$(curTotal, "### ### ### ###.00")
curSolde = Abs(xYBIACPT0.SOLDECEN)
libYBIASTO0_SOLDECEN = Format$(curSolde, "### ### ### ###.00")

If curTotal = curSolde Then
    libYBIASTO0_Diff = ""
Else
    libYBIASTO0_Diff.ForeColor = vbRed
    libYBIASTO0_Diff = Format$(curTotal - curSolde, "### ### ### ###.00")
End If

'xxxx Modification montant négatif 21/12/2009 Denis R.
Call fgYBIASTO0_Montant_Negatif
'xxxx FIN modification

'libYBIASTO0 = nbTotal & " dossiers, " & Format$(curTotal, "### ### ### ###.00") & " " & xYBIACPT0.COMPTEDEV
fgYBIASTO0.Visible = True
fgYBIASTO0_Sort1 = -1
End Sub

Public Sub fgYBIASTO0_DisplayLine(lIndex As Long)
On Error Resume Next
If chkSelect_YSTOMON = "0" And xYBIASTO0.YSTOMON = 0 Then Exit Sub
fgYBIASTO0.Rows = fgYBIASTO0.Rows + 1
fgYBIASTO0.Row = fgYBIASTO0.Rows - 1
fgYBIASTO0.Col = 0: fgYBIASTO0.Text = xYBIASTO0.YSTOAPP & " " & xYBIASTO0.YSTOOPE & " " & xYBIASTO0.YSTONAT & " " & xYBIASTO0.YSTONUM
fgYBIASTO0.Col = 1: fgYBIASTO0.Text = Format$(xYBIASTO0.YSTOMON, "### ### ### ###.00")
fgYBIASTO0.Col = 2: fgYBIASTO0.Text = xYBIASTO0.YSTODEV
fgYBIASTO0.Col = 3: fgYBIASTO0.Text = dateImp10(xYBIASTO0.YSTODEB)
fgYBIASTO0.Col = 4: fgYBIASTO0.Text = dateImp10(xYBIASTO0.YSTOFIN)
fgYBIASTO0.Col = 5: fgYBIASTO0.Text = xYBIASTO0.YSTOCLI
fgYBIASTO0.Col = 6: fgYBIASTO0.Text = Trim(xYBIASTO0.YSTOPCI)
fgYBIASTO0.Col = fgYBIASTO0_arrIndex: fgYBIASTO0.Text = lIndex
End Sub


Public Sub fgYBIASTO0_Reset()
fgYBIASTO0.Clear
fgYBIASTO0_Sort1 = 0: fgYBIASTO0_Sort2 = 0
fgYBIASTO0_Sort1_Old = -1
fgYBIASTO0_RowDisplay = 0: fgYBIASTO0_RowClick = 0
fgYBIASTO0_arrIndex = fgYBIASTO0.Cols - 1
blnfgYBIASTO0_DisplayLine = False
End Sub


Public Sub fgYBIASTO0_Sort()
If fgYBIASTO0.Rows > 1 Then
    fgYBIASTO0.Row = 1
    fgYBIASTO0.RowSel = fgYBIASTO0.Rows - 1
    
    If fgYBIASTO0_Sort1_Old = fgYBIASTO0_Sort1 Then
        If fgYBIASTO0_SortAD = 5 Then
            fgYBIASTO0_SortAD = 6
        Else
            fgYBIASTO0_SortAD = 5
        End If
    Else
        fgYBIASTO0_SortAD = 5
    End If
    fgYBIASTO0_Sort1_Old = fgYBIASTO0_Sort1
    
    fgYBIASTO0.Col = fgYBIASTO0_Sort1
    fgYBIASTO0.ColSel = fgYBIASTO0_Sort2
    fgYBIASTO0.Sort = fgYBIASTO0_SortAD
End If
'cboDevise_Reset
End Sub

Public Sub fgSelect_DisplayLine(lIndex As Integer)
On Error Resume Next

Dim curX As Currency
curX1 = Abs(xYBIACPT0.SOLDECEN)
curX2 = Abs(xYBIASTO0.YSTOMON)
curX = Abs(curX1 - curX2)

If chkSelect_Ecart.Value = "1" And curX = 0 Then Exit Sub
'$JPL 20101129
'curX1 = xYBIACPT0.SOLDECEN
'curX2 = xYBIASTO0.YSTOMON
'curX = curX1 - curX2

fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1
fgSelect.Col = 0: fgSelect.Text = xYBIASTO0.YSTOCLI & Asc10_13 & xYBIACPT0.COMPTECOM
fgSelect.Col = 1: fgSelect.Text = xYBIACPT0.CLIENARA1 & Asc10_13 & xYBIACPT0.COMPTEINT
fgSelect.Col = 2: fgSelect.Text = xYBIASTO0.YSTODEV

fgSelect.Col = 3: fgSelect.Text = Format$(curX1, "### ### ### ###.00") & Asc10_13 & Format$(curX2, "### ### ### ###.00")

''fgSelect.Col = 4: fgSelect.Text = Format$(xYBIASTO0.YSTOMON, "### ### ### ###.00")

fgSelect.Col = 4: fgSelect.CellForeColor = vbRed
If Trim(xYBIACPT0.COMPTECOM) <> "" Then
    If curX = 0 Then
        fgSelect.CellForeColor = vbBlue
        fgSelect.Text = ""
    Else
        fgSelect.Text = Format$(curX, "### ### ### ###.00")
    End If
    
Else
    fgSelect.Text = "? compte inconnu"
End If
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
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
    X = Format$(Val(fgSelect.Text), "000000000000000.00")
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub


Public Sub fgYBIASTO0_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgYBIASTO0.Rows - 1
    fgYBIASTO0.Row = I
    fgYBIASTO0.Col = lK
    X = Format$(Val(fgYBIASTO0.Text), "000000000000000.00")
    fgYBIASTO0.Col = fgYBIASTO0_arrIndex - 1
    fgYBIASTO0.Text = X
Next I


fgYBIASTO0_Sort1 = fgYBIASTO0_arrIndex - 1: fgYBIASTO0_Sort2 = fgYBIASTO0_arrIndex - 1
fgYBIASTO0_Sort
End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Private Sub mnuSelect_Print_Détail_Click_xlsManual(wsheet As Excel.Worksheet)
Me.Enabled = False: Me.MousePointer = vbHourglass
Call cmdPrint_Ok_xlsManual("D ", wsheet)
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuSelect_Print_Liste_Click_xlsManual(wsheet As Excel.Worksheet)
Me.Enabled = False: Me.MousePointer = vbHourglass
Call cmdPrint_Ok_xlsManual("L ", wsheet)
Me.Enabled = True: Me.MousePointer = 0
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), SAB_Stock_Aut)
'!!!sAB_Stock_Aut.Valider : consultation DAT
'blnSetfocus = True
Form_Init
Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case "@SAB_STOCK": 'blnAuto = True
                    If xlsManual Then
                        Call init_xlsManual
                        appExcelPublic.Workbooks.Add
                        Set wbExcel = appExcelPublic.ActiveWorkbook
                        With wbExcel
                            .Title = "BAL_Stock"
                            .Subject = "BAL_Stock"
                        End With
                    End If
                    chkSelect_Ecart = 1
                    cmdSelect_SQL
                    Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S60", "BIA-SAB-Stock-Ecart-Liste", "Archive")
                    If xlsManual Then
                        wbExcel.Sheets(1).Name = "BAL_Stock"
                        Call mnuSelect_Print_Liste_Click_xlsManual(wbExcel.Sheets(1))
                    Else
                        Call mnuSelect_Print_Liste_Click
                    End If
                    Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S60", "BIA-SAB-Stock-Ecart-Détail", "Archive")
                    If xlsManual Then
                        wbExcel.Sheets(2).Name = "BAL_Stock_detail"
                        Call mnuSelect_Print_Détail_Click_xlsManual(wbExcel.Sheets(2))
                    Else
                        mnuSelect_Print_Détail_Click
                    End If
                    chkSelect_Ecart = 0
                    cmdSelect_SQL
                    Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S60", "BIA-SAB-Stock-Liste", "Archive")
                    If xlsManual Then
                        wbExcel.Sheets(3).Name = "BAL_Stock_ecart0"
                        Call mnuSelect_Print_Liste_Click_xlsManual(wbExcel.Sheets(3))
                    Else
                        mnuSelect_Print_Liste_Click
                    End If
                    Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S60", "BIA-SAB-Stock-Détail", "Archive")
                    If xlsManual Then
                        wbExcel.Sheets.Add After:=Worksheets(3), Count:=1
                        wbExcel.Sheets(4).Name = "BAL_Stock_detail_ecart0"
                        Call mnuSelect_Print_Détail_Click_xlsManual(wbExcel.Sheets(4))
                    Else
                        mnuSelect_Print_Détail_Click
                    End If
                    
                     If xlsManual Then
                        Call wbExcel.Close(True)
                        Set wbExcel = Nothing
                    End If
                    Unload Me

    Case Else: 'blnAuto = False
End Select



End Sub

Public Sub Form_Init()
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistant", vbCritical, "frmYBIASTO0.param_init"
    Unload Me
Else
    lstErr.Clear
End If

blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fgYBIASTO0_FormatString = fgYBIASTO0.FormatString
fgSelect.Enabled = True
cmdReset
Me.Enabled = True
Me.MousePointer = 0
End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
currentAction = ""
rsYBIASTO0_Init meYBIASTO0
xYBIASTO0 = meYBIASTO0
fraSelect_Options.Enabled = False
SSTab1.Tab = 0
cmdSelect_Ok_Click
blnControl = True



End Sub


Public Function param_Init()

param_Init = Null
Call lstErr_Clear(lstErr, cmdContext, ". SAb_sTOCK_Import cbo"): DoEvents

fgSelect.Visible = False


Call rsYBIATAB0_cboK2("DEVISE", "ISO", cboSelect_YSTODEV)
rsZPLAN0_cboPLANCOOBL cboSelect_YSTOPCI

cboSelect_YSTOAPP.AddItem "   "
cboSelect_YSTOAPP.AddItem "CAU"
cboSelect_YSTOAPP.AddItem "CHG"
cboSelect_YSTOAPP.AddItem "CDO"
cboSelect_YSTOAPP.AddItem "CRE"
cboSelect_YSTOAPP.AddItem "DAT"
cboSelect_YSTOAPP.AddItem "ENC"
cboSelect_YSTOAPP.AddItem "TRE"
cboSelect_YSTOAPP.ListIndex = 0

cboSelect_YSTOOPE.AddItem "   "
cboSelect_YSTOOPE.AddItem "CDE"
cboSelect_YSTOOPE.AddItem "CDI"
cboSelect_YSTOOPE.AddItem "CHQ"
cboSelect_YSTOOPE.AddItem "EFF"
cboSelect_YSTOOPE.AddItem "EM1"
cboSelect_YSTOOPE.AddItem "EMP"
cboSelect_YSTOOPE.AddItem "ENG"
cboSelect_YSTOOPE.AddItem "GAR"
cboSelect_YSTOOPE.AddItem "PAR"
cboSelect_YSTOOPE.AddItem "PRE"
cboSelect_YSTOOPE.AddItem "RDE"
cboSelect_YSTOOPE.AddItem "RDI"
cboSelect_YSTOOPE.ListIndex = 0

cboSelect_YSTONAT.AddItem "   "
cboSelect_YSTONAT.ListIndex = 0

cboSelect_YSTOCLI.AddItem "   "
cboSelect_YSTOCLI.ListIndex = 0

fgSelect.Visible = True

Call lstErr_ChangeLastItem(lstErr, cmdContext, "= SAb_  Stock_Import"): DoEvents

Me.Enabled = True: Me.MousePointer = 0



End Function





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

Private Sub zoneImpression_xlsManual(lFct As String, nbRows As Long, wsheet As Excel.Worksheet)

    Call init_TypePagesetup
    If nbRows > 0 Then
        If Trim(lFct) <> "D" Then
            wsheet.Activate
            wsheet.Range("A1:F" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$F$" & CStr(nbRows)
            zoneImpressionPagesetup.LeftFooter = "&""Arial,Normal""&6&K04-024" & String(166, "_") & Chr(10) & "prtSAB_Stock   &D &T  BIA_INFO"
            zoneImpressionPagesetup.RightFooter = "&""Arial,Normal""&6&K04-024&P"
        Else
            wsheet.Activate
            wsheet.Range("A1:I" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$I$" & CStr(nbRows)
            zoneImpressionPagesetup.LeftFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "prtSAB_Stock   &D &T  BIA_INFO"
            zoneImpressionPagesetup.RightFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "&P"
            zoneImpressionPagesetup.Orientation = xlLandscape
            zoneImpressionPagesetup.Zoom = 95
        End If
    End If
    Call SetTypePageSetup(wsheet)

End Sub

Private Sub cboSelect_YSTOAPP_GotFocus()
txt_GotFocus cboSelect_YSTOAPP

End Sub

Private Sub cboSelect_YSTOAPP_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub cboSelect_YSTOAPP_LostFocus()
txt_LostFocus cboSelect_YSTOAPP

End Sub

Private Sub cboSelect_YSTOCLI_GotFocus()
txt_GotFocus cboSelect_YSTOCLI

End Sub

Private Sub cboSelect_YSTOCLI_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub cboSelect_YSTOCLI_LostFocus()
txt_LostFocus cboSelect_YSTOCLI

End Sub

Private Sub cboSelect_YSTODEV_GotFocus()
txt_GotFocus cboSelect_YSTODEV

End Sub

Private Sub cboSelect_YSTODEV_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub cboSelect_YSTODEV_LostFocus()
txt_LostFocus cboSelect_YSTODEV

End Sub

Private Sub cboSelect_YSTONAT_GotFocus()
txt_GotFocus cboSelect_YSTONAT

End Sub

Private Sub cboSelect_YSTONAT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub cboSelect_YSTONAT_LostFocus()
txt_LostFocus cboSelect_YSTONAT

End Sub

Private Sub cboSelect_YSTOOPE_GotFocus()
txt_GotFocus cboSelect_YSTOOPE

End Sub

Private Sub cboSelect_YSTOOPE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub cboSelect_YSTOOPE_LostFocus()
txt_LostFocus cboSelect_YSTOOPE

End Sub

Private Sub cboSelect_YSTOPCI_GotFocus()
txt_GotFocus cboSelect_YSTOPCI

End Sub

Private Sub cboSelect_YSTOPCI_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub cboSelect_YSTOPCI_LostFocus()
txt_LostFocus cboSelect_YSTOPCI

End Sub

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case SSTab1.Tab
    Case 0:
            If fgSelect.Rows > 1 Then
                Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
           End If
    Case 1:
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_SQL()
Dim xSQL As String, xWhere As String, xAnd As String
Dim blnOk As Boolean, blnCumul As Boolean
Dim I As Integer, nbDossier As Long

blnOk = False
nbDossier = 0

xWhere = ""
X = Trim(cboSelect_YSTOAPP)

If X <> "" Then
    If xWhere = "" Then
        xAnd = " where "
    Else
        xAnd = " and "
    End If
    xWhere = xWhere & xAnd & "YSTOAPP = '" & X & "'"
End If
X = Trim(cboSelect_YSTOOPE)
If X <> "" Then
    If xWhere = "" Then
        xAnd = " where "
    Else
        xAnd = " and "
    End If
    xWhere = xWhere & xAnd & "YSTOOPE = '" & X & "'"
End If
X = Trim(cboSelect_YSTONAT)
If X <> "" Then
    If xWhere = "" Then
        xAnd = " where "
    Else
        xAnd = " and "
    End If
    xWhere = xWhere & xAnd & "YSTONAT = '" & X & "'"
End If
X = Trim(cboSelect_YSTOPCI)
If X <> "" Then
    If xWhere = "" Then
        xAnd = " where "
    Else
        xAnd = " and "
    End If
    xWhere = xWhere & xAnd & "YSTOPCI like '" & X & "%'"
End If
X = Trim(cboSelect_YSTOCLI)
If X <> "" Then
    If xWhere = "" Then
        xAnd = " where "
    Else
        xAnd = " and "
    End If
    xWhere = xWhere & xAnd & "YSTOCLI = " & X
End If
X = Trim(cboSelect_YSTODEV)
If X <> "" Then
    If xWhere = "" Then
        xAnd = " where "
    Else
        xAnd = " and "
    End If
    xWhere = xWhere & xAnd & "YSTODEV = '" & X & "'"
End If

Call YBIASTO0_Sql(xWhere, nbDossier, arrYBIASTO0(), arrYBIACPT0(), arrCompte_Nb)

Call lstErr_AddItem(lstErr, cmdContext, "Lignes d'encours : " & nbDossier): DoEvents

fgSelect_Display

End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAb_Stock_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
fgYBIASTO0.Clear
If blnOk Then
    cmdSelect_Ok.Caption = "Options"
    cmdSelect_Ok.BackColor = &HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = False
    cmdSelect_SQL
Else
    cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAb_Stock_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
        xYBIASTO0 = arrYBIASTO0(K)
        xYBIACPT0 = arrYBIACPT0(K)
        fgYBIASTO0_Display
        SSTab1.Tab = 1
        
   End If
End If
End Sub

Private Sub fgYBIASTO0_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
On Error Resume Next
If y <= fgYBIASTO0.RowHeightMin Then
    Select Case fgYBIASTO0.Col
        Case 0: fgYBIASTO0_Sort1 = 0: fgYBIASTO0_Sort2 = 1: fgYBIASTO0_Sort
        Case 1:  fgYBIASTO0_Sort1 = 1: fgYBIASTO0_Sort2 = 1: fgYBIASTO0_SortX 1
        Case 2: fgYBIASTO0_Sort1 = 2: fgYBIASTO0_Sort2 = 2: fgYBIASTO0_Sort
        Case 3: fgYBIASTO0_Sort1 = 3: fgYBIASTO0_Sort2 = 3: fgYBIASTO0_Sort
        Case 4: fgYBIASTO0_Sort1 = 4: fgYBIASTO0_Sort2 = 4: fgYBIASTO0_Sort
    End Select
Else
    If fgYBIASTO0.Rows > 1 Then
        Call fgYBIASTO0_Color(fgYBIASTO0_RowClick, MouseMoveUsr.BackColor, fgYBIASTO0_ColorClick)
        fgYBIASTO0.Col = fgYBIASTO0_arrIndex:  K = CLng(fgYBIASTO0.Text)
        'xYBIASTO0 = selYBIASTO0(K)
        'srvYBIASTO0_ElpDisplay xYBIASTO0

   End If
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
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


If SSTab1.Tab = 0 Then
        Unload Me
    Exit Sub
Else
    SSTab1.Tab = SSTab1.Tab - 1
End If

End Sub

Public Sub cmdContext_Return()
On Error Resume Next
If SSTab1.Tab = 0 Then
    fgSelect.Row = fgSelect.TopRow
    fgSelect.Col = fgSelect_arrIndex: ' wK1 = fgSelect.Text
    'cmdSelect txtSelect ''fgSelect.Text

'    cmdSelect_Click
Else
    SendKeys "{TAB}"
End If
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
On Error GoTo Error_Handler

mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False

Exit Sub

Error_Handler:
blnControl = False
MsgBox Error
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


Public Sub txt_X()
'Call txt_GotFocus(txt)
'KeyAscii = convUCase(KeyAscii)
'Call txt_LostFocus(txt)

'Call txt_GotFocus(txt)
'If XopDevise(2).maxD = 0 Then
'    Call num_KeyAscii(KeyAscii)
'Else
'    Call num_KeyAsciiD(KeyAscii, txt)
'End If
'Call txt_LostFocus(txt)
'Call txt_GotFocus(txt)
'Call txt_LostFocus(txt)

End Sub





Private Sub mnuSelect_Print_Détail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Ok "D "
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_Liste_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Ok "L "
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then cmdSelect_Ok.SetFocus

End Sub














Private Sub SSTab1_GotFocus()
Select Case SSTab1.Tab
    Case 0: fgSelect.LeftCol = 0
    Case 1: fgYBIASTO0.LeftCol = 0
End Select
End Sub


Public Sub cmdPrint_Ok(lFct As String)
fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Liste : " & fgSelect.Rows - 1)

prtSAB_Stock_Monitor lFct, fgSelect, arrYBIASTO0(), arrYBIACPT0(), arrCompte_Nb
fgSelect.Visible = True
Me.Show
End Sub
