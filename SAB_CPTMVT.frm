VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSAB_CPTMVT 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_CPTMVT : Analyse des mouvements comptables"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "SAB_CPTMVT.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   45
      Width           =   5055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Sélection"
      TabPicture(0)   =   "SAB_CPTMVT.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "?"
      TabPicture(1)   =   "SAB_CPTMVT.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraSelect 
         Height          =   8445
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13560
         Begin VB.CheckBox chkSelect_All 
            Caption         =   "Tous les comptes"
            Height          =   495
            Left            =   3840
            TabIndex        =   14
            Top             =   195
            Width           =   1695
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect_D 
            Height          =   7455
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Visible         =   0   'False
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   13150
            _Version        =   393216
            Cols            =   4
            BackColorFixed  =   8438015
            BackColorBkg    =   12640511
            GridColor       =   255
            AllowUserResizing=   1
            FormatString    =   $"SAB_CPTMVT.frx":0044
         End
         Begin MSFlexGridLib.MSFlexGrid fgDossier 
            Height          =   7485
            Left            =   6360
            TabIndex        =   13
            Top             =   720
            Width           =   7200
            _ExtentX        =   12700
            _ExtentY        =   13203
            _Version        =   393216
            Rows            =   1
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   15269886
            ForeColor       =   8388608
            BackColorFixed  =   12648447
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   15269886
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"SAB_CPTMVT.frx":0120
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CheckBox chkSelect 
            Caption         =   "Afficher les dossiers soldés"
            Height          =   255
            Left            =   11160
            TabIndex        =   11
            Top             =   320
            Width           =   2295
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7545
            Left            =   2880
            TabIndex        =   8
            Top             =   720
            Width           =   10560
            _ExtentX        =   18627
            _ExtentY        =   13309
            _Version        =   393216
            Rows            =   1
            Cols            =   12
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   16777210
            ForeColor       =   8388608
            BackColorFixed  =   16776921
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   16777210
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   "<Compte             |>Dossier     |>Solde DB                  |>Solde CR              |>Date TRT   ||"
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
         Begin VB.ListBox lstSelect 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7260
            Left            =   120
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   10
            Top             =   720
            Width           =   6500
         End
         Begin VB.ComboBox cboSelect_SQL 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   9
            Text            =   "cboSelect_SQL"
            Top             =   260
            Width           =   3615
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   405
            Left            =   7800
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   260
            Width           =   3015
         End
         Begin VB.Frame fraSelect_Options 
            Height          =   285
            Left            =   12720
            TabIndex        =   6
            Top             =   120
            Width           =   795
         End
         Begin MSComCtl2.DTPicker txtSelect_AmjMin 
            Height          =   300
            Left            =   5760
            TabIndex        =   16
            Top             =   320
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CheckBox        =   -1  'True
            CustomFormat    =   "dd  MM yyy"
            Format          =   122028035
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin VB.Label lblSelect_Param 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1335
            Left            =   7000
            TabIndex        =   15
            Top             =   720
            Width           =   6000
         End
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "SAB_CPTMVT.frx":01E1
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label libRéférenceInterne 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuContext_x1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuselect 
      Caption         =   "mnuSelect"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Quit 
         Caption         =   "Abandonner"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint0_All 
         Caption         =   "Imprimer TOUS les courriers"
      End
   End
End
Attribute VB_Name = "frmSAB_CPTMVT"
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
Dim SAB_CDRAut As typeAuthorization
Dim blnTransaction As Boolean
Dim blnAuto As Boolean, blnAuto_Ok As Boolean
Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long
Dim wAmjMin7 As Long, wAmjMax7 As Long


Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnSetfocus As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim fgSelect_D_FormatString As String

Dim fgDossier_FormatString As String, fgDossier_K As Integer
Dim fgDossier_RowDisplay As Integer, fgDossier_RowClick As Integer, fgDossier_ColClick As Integer
Dim fgDossier_ColorClick As Long, fgDossier_ColorDisplay As Long
Dim fgDossier_Sort1 As Integer, fgDossier_Sort2 As Integer
Dim fgDossier_SortAD As Integer, fgDossier_Sort1_Old As Integer
Dim fgDossier_arrIndex As Integer
Dim blnfgDossier_DisplayLine As Boolean

'______________________________________________________________________

Dim xZMOUVEA0 As typeZMOUVEA0
Dim arrZMOUVEA0() As typeZMOUVEA0, arrZMOUVEA0_Nb As Long, arrZMOUVEA0_Max As Long, arrZMOUVEA0_Index As Long
Dim selZMOUVEA0() As typeZMOUVEA0, selZMOUVEA0_Nb As Long, selZMOUVEA0_Max As Long, selZMOUVEA0_Index As Long
Dim selYBIACPT0() As typeYBIACPT0, selYBIACPT0_Nb As Long, selYBIACPT0_Max As Long, selYBIACPT0_Index As Long
Dim xYBIACPT0 As typeYBIACPT0
Dim mMOUVEMCOM As String

Dim xYBIAMVT0 As typeYBIAMVT0
Dim cmdSelect_Ok_Caption As String
Dim cmdSelect_SQL_K As String

Dim curDB As Currency, curCR As Currency
Dim selZCHGOPE0() As typeZCHGOPE0, selZCHGOPE0_Nb As Long, selZCHGOPE0_Max As Long, selZCHGOPE0_Index As Long
Dim xZCHGOPE0 As typeZCHGOPE0
Public Sub fgDossier_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgDossier.Row

If lRow > 0 And lRow < fgDossier.Rows Then
    fgDossier.Row = lRow
    For I = 0 To fgDossier_arrIndex
        fgDossier.Col = I: fgDossier.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgDossier.Row = mRow
    If fgDossier.Row > 0 Then
        lRow = fgDossier.Row
        lColor_Old = fgDossier.CellBackColor
        For I = fgDossier_arrIndex To 0 Step -1
          fgDossier.Col = I: fgDossier.CellBackColor = lColor
        Next I
        fgDossier.LeftCol = 0
    End If
End If

End Sub

Private Sub fgDossier_Display()
Dim K As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgDossier_Reset
fgDossier.Visible = False: Me.MousePointer = vbHourglass: DoEvents
cmdPrint.Enabled = False

fgDossier.Rows = 1
fgDossier.FormatString = fgDossier_FormatString
currentAction = "fgdossier_Display"
libRéférenceInterne = "Compte = " & xZMOUVEA0.MOUVEMCOM & " Dossier = " & xZMOUVEA0.MOUVEMNUM
Call lstErr_Clear(lstErr, cmdContext, libRéférenceInterne)
Call lstErr_AddItem(lstErr, cmdContext, "Solde = " & Format$(xZMOUVEA0.MOUVEMMON, "### ### ### ###.00"))
DoEvents
xWhere = "Where MOUVEMCOM = '" & xZMOUVEA0.MOUVEMCOM & "' and MOUVEMNUM = " & xZMOUVEA0.MOUVEMNUM
arrZMOUVEA0_SQL xWhere
For K = 1 To arrZMOUVEA0_Nb
    xZMOUVEA0 = arrZMOUVEA0(K)
        fgDossier.Rows = fgDossier.Rows + 1
        fgDossier.Row = fgDossier.Rows - 1
        fgDossier_DisplayLine K
Next K

Call lstErr_AddItem(lstErr, cmdContext, "Nb Mvts : " & arrZMOUVEA0_Nb): DoEvents
'If fgDossier.Rows > 1 Then
'    fgDossier_Sort1 = 0: fgDossier_Sort2 = 1: fgDossier_Sort
'    cmdPrint.Enabled = True
'End If
fgDossier.Visible = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgDossier_Display_2()
Dim K As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgDossier_Reset
fgDossier.Visible = False: Me.MousePointer = vbHourglass: DoEvents
cmdPrint.Enabled = False

fgDossier.Rows = 1
fgDossier.FormatString = fgDossier_FormatString
currentAction = "fgdossier_Display_2"
libRéférenceInterne = "Compte = " & xYBIACPT0.COMPTECOM & " Dossier = " & xYBIACPT0.COMPTEINT
Call lstErr_Clear(lstErr, cmdContext, libRéférenceInterne)
Call lstErr_AddItem(lstErr, cmdContext, "Solde = " & Format$(xYBIACPT0.SOLDECEN, "### ### ### ###.00"))
DoEvents
xWhere = "Where MOUVEMCOM = '" & xYBIACPT0.COMPTECOM & "' "
arrZMOUVEA0_SQL xWhere
For K = 1 To arrZMOUVEA0_Nb
    xZMOUVEA0 = arrZMOUVEA0(K)
        fgDossier.Rows = fgDossier.Rows + 1
        fgDossier.Row = fgDossier.Rows - 1
        fgDossier_DisplayLine K
Next K

Call lstErr_AddItem(lstErr, cmdContext, "Nb Mvts : " & arrZMOUVEA0_Nb): DoEvents
fgDossier.Visible = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgDossier_Display_3()
Dim K As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgDossier_Reset
fgDossier.Visible = False: Me.MousePointer = vbHourglass: DoEvents
cmdPrint.Enabled = False

fgDossier.Rows = 1
fgDossier.FormatString = fgDossier_FormatString
currentAction = "fgdossier_Display_3"
libRéférenceInterne = " Dossier = " & xZCHGOPE0.CHGOPEDOS
Call lstErr_Clear(lstErr, cmdContext, libRéférenceInterne)
DoEvents
xWhere = "Where MOUVEMNUM = " & xZCHGOPE0.CHGOPEDOS & " and MOUVEMOPE = '" & xZCHGOPE0.CHGOPEOPE & "' "
arrZMOUVEA0_SQL xWhere
For K = 1 To arrZMOUVEA0_Nb
    xZMOUVEA0 = arrZMOUVEA0(K)
        fgDossier.Rows = fgDossier.Rows + 1
        fgDossier.Row = fgDossier.Rows - 1
        fgDossier_DisplayLine K
Next K

Call lstErr_AddItem(lstErr, cmdContext, "Nb Mvts : " & arrZMOUVEA0_Nb): DoEvents
fgDossier.Visible = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub fgDossier_DisplayLine(lIndex As Long)
On Error Resume Next
fgDossier.Col = 0: fgDossier.Text = dateIBM10(xZMOUVEA0.MOUVEMDTR, True)
fgDossier.Col = 1: fgDossier.Text = xZMOUVEA0.MOUVEMOPE & " " & xZMOUVEA0.MOUVEMEVE & " " & Trim(xZMOUVEA0.MOUVEMANA)
fgDossier.Col = 2: fgDossier.Text = Format$(xZMOUVEA0.MOUVEMMON, "### ### ### ###.00")
If xZMOUVEA0.MOUVEMMON > 0 Then
    fgDossier.CellForeColor = vbRed
Else
    fgDossier.CellForeColor = vbBlue
End If

fgDossier.Col = 3: fgDossier.Text = xZMOUVEA0.MOUVEMCOM
fgDossier.Col = 4: fgDossier.Text = xZMOUVEA0.MOUVEMNUM
fgDossier.Col = 5: fgDossier.Text = xZMOUVEA0.MOUVEMPIE & " " & xZMOUVEA0.MOUVEMECR
fgDossier.Col = fgDossier_arrIndex: fgDossier.Text = lIndex

End Sub


Public Sub fgDossier_Reset()
fgDossier.Clear
fgDossier_Sort1 = 0: fgDossier_Sort2 = 0
fgDossier_Sort1_Old = -1
fgDossier_RowDisplay = 0: fgDossier_RowClick = 0
fgDossier_arrIndex = fgDossier.Cols - 1
blnfgDossier_DisplayLine = False
End Sub


Public Sub fgDossier_Sort()
If fgDossier.Rows > 1 Then
    fgDossier.Row = 1
    fgDossier.RowSel = fgDossier.Rows - 1
    
    If fgDossier_Sort1_Old = fgDossier_Sort1 Then
        If fgDossier_SortAD = 5 Then
            fgDossier_SortAD = 6
        Else
            fgDossier_SortAD = 5
        End If
    Else
        fgDossier_SortAD = 5
    End If
    fgDossier_Sort1_Old = fgDossier_Sort1
    
    fgDossier.Col = fgDossier_Sort1
    fgDossier.ColSel = fgDossier_Sort2
    fgDossier.Sort = fgDossier_SortAD
End If
'cboDevise_Reset
End Sub


'______________________________________________________________________

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
        For I = fgSelect_arrIndex To 0 Step -1
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.LeftCol = 0
    End If
End If

End Sub
Private Sub fgSelect_Display()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
cmdPrint.Enabled = False
currentAction = "fgselect_Display"

For I = 1 To selZMOUVEA0_Nb

    If chkSelect = "0" And selZMOUVEA0(I).MOUVEMMON = 0 Then
    Else
        xZMOUVEA0 = selZMOUVEA0(I)
    
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
    End If
Next I

Call lstErr_Clear(lstErr, cmdContext, "Nb enregistrements : " & fgSelect.Rows - 1): DoEvents
If fgSelect.Rows > 1 Then
'    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
    cmdPrint.Enabled = True
End If
fgSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub fgSelect_Display_2()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
fgSelect.Rows = 1
fgSelect.FormatString = "<Compte           |<Intitulé              |>Solde DB             |>Solde CR             |>Date Mvt "
cmdPrint.Enabled = False
currentAction = "fgselect_Display"

For I = 1 To selYBIACPT0_Nb

    If chkSelect = "0" And selYBIACPT0(I).SOLDECEN = 0 Then
    Else
        xYBIACPT0 = selYBIACPT0(I)
    
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine_2 I
    End If
Next I

Call lstErr_Clear(lstErr, cmdContext, "Nb enregistrements : " & fgSelect.Rows - 1): DoEvents
If fgSelect.Rows > 1 Then
'    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
    cmdPrint.Enabled = True
End If
fgSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub


Private Sub fgSelect_Display_3()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
fgSelect.Rows = 1
fgSelect.FormatString = "<OPE NAT      |>Dossier |<Sens|<Dev |>Montant         |<Contrepartie   |>Date création "
cmdPrint.Enabled = False
currentAction = "fgselect_Display"

For I = 1 To selZCHGOPE0_Nb

        xZCHGOPE0 = selZCHGOPE0(I)
    
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine_3 I
    
Next I

Call lstErr_Clear(lstErr, cmdContext, "Nb enregistrements : " & fgSelect.Rows - 1): DoEvents
If fgSelect.Rows > 1 Then
'    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
    cmdPrint.Enabled = True
End If
fgSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub



Private Sub lstSelect_Load_1()
Dim I As Long, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String
Dim X As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
lstSelect.Visible = False
lstSelect.Clear
currentAction = "lstSelect_Load_1"
xWhere = ""

Select Case cmdSelect_SQL_K
    Case "1":   xWhere = " where COMPTECOM like 'R911329%'"
    Case "4":   xWhere = " where COMPTECOM like 'R987520%'"
    Case "5":
        X = InputBox("saisir le PCI (par exemple 911329)", "Rapprochement Dossier", "911209")
        xWhere = " where COMPTECOM like 'R" & Val(X) & "%'"
End Select
Set rsSab = Nothing

xSQL = "select COMPTECOM,COMPTEINT,COMPTEDEV from " & paramIBM_Library_SAB & ".ZCOMPTE0 " & xWhere & " order by COMPTEDEV,COMPTECOM"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    lstSelect.AddItem rsSab("COMPTEDEV") & " " & rsSab("COMPTECOM") & " " & rsSab("COMPTEINT")
    rsSab.MoveNext

Loop
Call lstErr_Clear(lstErr, cmdContext, "Nb enregistrements : " & lstSelect.ListCount): DoEvents
lstSelect.Visible = True
lstSelect.Enabled = True
'lstSelect.SetFocus
If lstSelect.ListCount > 0 Then lstSelect.ListIndex = 0
cmdSelect_Ok_Caption = "Extraire les mouvements"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
chkSelect.Visible = True
chkSelect_All.Visible = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
Dim X As String

On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = xZMOUVEA0.MOUVEMCOM
fgSelect.Col = 1: fgSelect.Text = xZMOUVEA0.MOUVEMNUM
X = Format$(Abs(xZMOUVEA0.MOUVEMMON), "### ### ### ###.00")
If xZMOUVEA0.MOUVEMMON > 0 Then
    fgSelect.Col = 2: fgSelect.Text = X
    fgSelect.CellForeColor = vbRed
Else
    fgSelect.Col = 3: fgSelect.Text = X
    fgSelect.CellForeColor = vbBlue
End If
fgSelect.Col = 4: fgSelect.Text = dateIBM10(xZMOUVEA0.MOUVEMDTR, True)

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

End Sub

Public Sub fgSelect_DisplayLine_2(lIndex As Long)
Dim X As String

On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = xYBIACPT0.COMPTECOM
fgSelect.Col = 1: fgSelect.Text = xYBIACPT0.COMPTEINT
X = Format$(Abs(xYBIACPT0.SOLDECEN), "### ### ### ###.00")
If xYBIACPT0.SOLDECEN > 0 Then
    fgSelect.Col = 2: fgSelect.Text = X
    fgSelect.CellForeColor = vbRed
Else
    fgSelect.Col = 3: fgSelect.Text = X
    fgSelect.CellForeColor = vbBlue
End If
fgSelect.Col = 4: fgSelect.Text = dateIBM10(xYBIACPT0.SOLDEDMO, True)

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

End Sub



Public Sub fgSelect_DisplayLine_3(lIndex As Long)
Dim X As String

On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = xZCHGOPE0.CHGOPEOPE & " " & xZCHGOPE0.CHGOPENAT & " " & xZCHGOPE0.CHGOPESER & " " & xZCHGOPE0.CHGOPESSE
fgSelect.Col = 1: fgSelect.Text = xZCHGOPE0.CHGOPEDOS
fgSelect.Col = 2: fgSelect.Text = xZCHGOPE0.CHGOPESEN
fgSelect.Col = 3: fgSelect.Text = xZCHGOPE0.CHGOPEDE1
X = Format$(Abs(xZCHGOPE0.CHGOPEMO1), "### ### ### ###.00")
fgSelect.Col = 4: fgSelect.Text = X
fgSelect.Col = 5: fgSelect.Text = xZCHGOPE0.CHGOPECON
fgSelect.Col = 6: fgSelect.Text = dateIBM10(xZCHGOPE0.CHGOPECRE, True)

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
    If lK = 2 Then
        fgSelect.Col = 2
        X = fgSelect.Text
    Else
        X = ""
    End If
    
    fgSelect.Col = 3
    X = X & Format$(Val(fgSelect.Text), "000000000000000.00")
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


Public Sub cmdContext_Quit()
If fgSelect_D.Visible Then fgSelect_D.Visible = False: Exit Sub
If fgDossier.Visible Then libRéférenceInterne = "": fgDossier.Visible = False: Exit Sub
If fgSelect.Visible Then fgSelect.Visible = False: lstSelect.Enabled = True: cmdSelect_Ok.Caption = "Extraire les mouvements": Exit Sub
If lstSelect.Visible Then lstSelect.Visible = False: Exit Sub

Unload Me
End Sub




Private Sub cboSelect_SQL_Click()
cmdSelect_SQL_K = Mid$(cboSelect_SQL, 1, 1)
If blnControl Then
    chkSelect.Visible = False
    chkSelect_All.Visible = False
    chkSelect_All.Caption = "Tous les comptes"
    lblSelect_Param.Visible = False
    txtSelect_AMJMIN.Visible = False
    Me.Enabled = False: Me.MousePointer = vbHourglass
    
    Select Case cmdSelect_SQL_K
        Case "1": lstSelect_Load_1
        Case "2": lstSelect_Load_2
        Case "3": lstSelect_Load_3
        Case "4", "5": lstSelect_Load_1
   End Select
    Me.Enabled = True: Me.MousePointer = 0
End If
End Sub


Private Sub chkSelect_All_Click()
lstSelect_All

End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
Dim I As Integer

blnControl = False
usrColor_Set

cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
blncmdOk_Visible = False: blncmdSave_Visible = False
currentAction = ""

blnAuto = False
blnAuto_Ok = False
lstSelect.Visible = False
cmdSelect_Ok.Caption = "Extraire les mouvements"

libRéférenceInterne = ""
cboSelect_SQL.ListIndex = 0
fgSelect.Visible = False
fgDossier.Visible = False
fgSelect_D.Visible = False
chkSelect.Visible = False
chkSelect_All.Visible = False
lblSelect_Param.Visible = False
lblSelect_Param.ForeColor = vbMagenta
Call DTPicker_Set(txtSelect_AMJMIN, YBIATAB0_DATE_CPT_J)

blnControl = True
End Sub
Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

SSTab1.Tab = 0

blnControl = False

fgSelect_FormatString = fgSelect.FormatString
fgDossier_FormatString = fgDossier.FormatString
cmdSelect_Ok.Visible = False
fraSelect_Options.Visible = False
cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1 - CDO : R911329*"
cboSelect_SQL.AddItem "2 - CPT : Dettes & Créances rattachées"
cboSelect_SQL.AddItem "3 - TRF : Change / Transfert"
cboSelect_SQL.AddItem "4 - CDO : R987520*"
cboSelect_SQL.AddItem "5 - CDO : saisie manuelle du PCI"
cmdReset



End Sub

Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
         If TypeOf xobj Is CommandButton Or TypeOf xobj Is ListBox Then
           xobj.BackColor = MouseMoveActiveControl.BackColor
        Else
            xobj.ForeColor = MouseMoveActiveControl.ForeColor
        End If
        Exit For
    End If
Next xobj

End Sub

Public Sub MouseMoveActiveControl_Set(C As Control)
If MouseMoveActiveControl_Name <> C.Name Then
    MouseMoveActiveControl_Reset
    If Not C.Enabled Then
        MouseMoveActiveControl_Name = ""
    Else
        MouseMoveActiveControl_Name = C.Name
        If TypeOf C Is CommandButton Or TypeOf C Is ListBox Then
            
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            MouseMoveActiveControl.ForeColor = C.ForeColor
            C.ForeColor = MouseMoveUsr.ForeColor
        End If
    End If
End If

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

Private Sub cmdPrint_Click()
Dim Msg As String
Dim I As Integer

Me.Enabled = False: Me.MousePointer = vbHourglass
    Select Case cmdSelect_SQL_K
        Case "1", "4", "5": cmdPrint_Ok_1
        Case "2": cmdPrint_Ok_2
    End Select

Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long


blnOk = lstSelect.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAB_CDR_cmdSelect_Ok ........"): DoEvents
cmdSelect_Ok.Visible = False
fgSelect.Clear
DoEvents
If blnOk Then
    cmdSelect_Ok.Caption = "Choisir les comptes"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    lstSelect.BackColor = &H8000000F
    Call usrColor_Container(lstSelect, lstSelect.BackColor)
    lstSelect.Enabled = False
    Select Case cmdSelect_SQL_K
        Case "1", "4", "5": cmdSelect_SQL
        Case "2": cmdSelect_SQL_2
        Case "3": cmdSelect_SQL_3
    End Select

    lstSelect.Enabled = False: chkSelect_All.Enabled = False
Else
    cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
    cmdSelect_Ok.BackColor = &HC0FFC0
    lstSelect.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(lstSelect, lstSelect.BackColor)
    If cmdSelect_SQL_K = "1" Then lstSelect_All
    lstSelect.Enabled = True
    fgSelect.Visible = False
    lstSelect.Enabled = True: chkSelect_All.Enabled = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_CDR_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
cmdSelect_Ok.Visible = True

End Sub


Private Sub cmdSelect_SQL()
Dim V
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL"): DoEvents

currentAction = "cmdSelect_SQL"
xWhere = ""
ReDim selZMOUVEA0(501)
selZMOUVEA0_Max = 500: selZMOUVEA0_Nb = 0
rsZMOUVEA0_Init selZMOUVEA0(0)
selZMOUVEA0(0).MOUVEMNUM = -1

Set rsSab = Nothing
If chkSelect_All = "1" Then
    xSQL = "select MOUVEMCOM,MOUVEMNUM,MOUVEMMON,MOUVEMDTR from " & paramIBM_Library_SAB & ".ZMOUVEA0  where MOUVEMCOM like 'R911329%' order by MOUVEMCOM,MOUVEMNUM,MOUVEMDTR"
    Set rsSab = cnsab.Execute(xSQL)
    cmdSelect_SQL_1
    xSQL = "select MOUVEMCOM,MOUVEMNUM,MOUVEMMON from " & paramIBM_Library_SABSPE & ".YCDOR911 "
    Set rsSab = cnsab.Execute(xSQL)
    cmdSelect_SQL_1_Reprise
Else
    For K = 0 To lstSelect.ListCount - 1
        lstSelect.ListIndex = K
    '    If lstSelect.Selected(lstSelect.ListIndex) Then
        If lstSelect.Selected(K) Then
            Call lstErr_AddItem(lstErr, cmdContext, lstSelect.Text): DoEvents
    
            xWhere = " where MOUVEMCOM = '" & Mid$(lstSelect.Text, 5, 20) & "'"
            xSQL = "select MOUVEMCOM,MOUVEMNUM,MOUVEMMON,MOUVEMDTR from " & paramIBM_Library_SAB & ".ZMOUVEA0 " & xWhere & " order by MOUVEMNUM,MOUVEMDTR"
            Set rsSab = cnsab.Execute(xSQL)
            cmdSelect_SQL_1
             'xWhere = " where MOUVEMCOM = '" & Mid$(lstSelect.Text, 5, 20) & "'"
            xSQL = "select MOUVEMCOM,MOUVEMNUM,MOUVEMMON from " & paramIBM_Library_SABSPE & ".YCDOR911 " & xWhere
            Set rsSab = cnsab.Execute(xSQL)
            cmdSelect_SQL_1_Reprise
       End If

    Next K
End If
    
fgSelect_Display
If chkSelect_All = "1" Then chkSelect_All = "0"

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub cmdSelect_SQL_2()
Dim V
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_2"): DoEvents

currentAction = "cmdSelect_SQL_2"
xWhere = ""
ReDim selYBIACPT0(501)
selYBIACPT0_Max = 500: selYBIACPT0_Nb = 0
rsYBIACPT0_Init selYBIACPT0(0)

Set rsSab = Nothing
For K = 0 To lstSelect.ListCount - 1
    lstSelect.ListIndex = K
    If lstSelect.Selected(K) Then
        Call lstErr_AddItem(lstErr, cmdContext, lstSelect.Text): DoEvents

        xWhere = " where PLANCOPRO = '" & Mid$(lstSelect.Text, 1, 3) & "'"
        xSQL = "select COMPTECOM,COMPTEINT,SOLDECEN,SOLDEDMO from " & paramIBM_Library_SABSPE & ".YBIACPT0 " & xWhere & " order by COMPTECOM"
        Set rsSab = cnsab.Execute(xSQL)
        cmdSelect_SQL_2_C
    End If
Next K
    
fgSelect_Display_2
If chkSelect_All = "1" Then chkSelect_All = "0"

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_3()
Dim V
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_3"): DoEvents

currentAction = "cmdSelect_SQL_3"

xWhere = ""

Call DTPicker_Control(txtSelect_AMJMIN, wAMJMin)
If wAMJMin <> "00000000" Then xAnd = " and CHGOPECRE = " & wAMJMin - 19000000

ReDim selZCHGOPE0(501)
selZCHGOPE0_Max = 500: selZCHGOPE0_Nb = 0
rsZCHGOPE0_Init selZCHGOPE0(0)

Set rsSab = Nothing
'20060110 paramIBM_Library_SAB = "JPLTST"
'20060110 MsgBox "cmdSelect_SQL_3 : JPLTST"
If chkSelect_All = "1" And wAMJMin <> "00000000" Then
        xWhere = " where CHGOPECRE = " & wAMJMin - 19000000
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZCHGOPE0 " & xWhere & " order by CHGOPEDOS"
        Set rsSab = cnsab.Execute(xSQL)
        cmdSelect_SQL_3_C
Else
    For K = 0 To lstSelect.ListCount - 1
        lstSelect.ListIndex = K
        If lstSelect.Selected(K) Then
            Call lstErr_AddItem(lstErr, cmdContext, lstSelect.Text): DoEvents
    
            xWhere = " where CHGOPEOPE = '" & Mid$(lstSelect.Text, 1, 3) & "'" _
                   & " and CHGOPENAT = '" & Mid$(lstSelect.Text, 4, 3) & "'" & xAnd
            xSQL = "select * from " & paramIBM_Library_SAB & ".ZCHGOPE0 " & xWhere & " order by CHGOPEDOS"
            Set rsSab = cnsab.Execute(xSQL)
            cmdSelect_SQL_3_C
        End If
    Next K
End If

fgSelect_Display_3

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub arrZMOUVEA0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrZMOUVEA0(101)
arrZMOUVEA0_Max = 100: arrZMOUVEA0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SAB & ".ZMOUVEA0 " & xWhere & " order by MOUVEMDTR,MOUVEMPIE"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsZMOUVEA0_GetBuffer(rsSab, xZMOUVEA0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgselect_Display"
        '' Exit Sub
     Else
         arrZMOUVEA0_Nb = arrZMOUVEA0_Nb + 1
         If arrZMOUVEA0_Nb > arrZMOUVEA0_Max Then
             arrZMOUVEA0_Max = arrZMOUVEA0_Max + 50
             ReDim Preserve arrZMOUVEA0(arrZMOUVEA0_Max)
         End If
         
         arrZMOUVEA0(arrZMOUVEA0_Nb) = xZMOUVEA0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub fgDossier_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
Dim xWhere As String, xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass
On Error Resume Next
If y <= fgDossier.RowHeightMin Then
    Select Case fgDossier.Col
        Case 0: fgDossier_Sort1 = 0: fgDossier_Sort2 = 1: fgDossier_Sort
        Case 1:  fgDossier_Sort1 = 1: fgDossier_Sort2 = 1: fgDossier_Sort
        Case 2: fgDossier_Sort1 = 2: fgDossier_Sort2 = 2: fgDossier_Sort
        Case 3: fgDossier_Sort1 = 3: fgDossier_Sort2 = 3: fgDossier_Sort
       'Case fgDossier_arrIndex:  fgDossier_SortX fgDossier_arrIndex
    End Select
Else
    If fgDossier.Rows > 1 Then
         fgDossier.Col = fgDossier_arrIndex:  K = CLng(fgDossier.Text)
       Call fgDossier_Color(fgDossier_RowClick, MouseMoveUsr.BackColor, fgDossier_ColorClick)
        
        xZMOUVEA0 = arrZMOUVEA0(K)
        If cmdSelect_SQL_K = "3" Then
            fgSelect_D_ZLIBEL0
        Else
            xWhere = "Where MOUVEMPIE = '" & xZMOUVEA0.MOUVEMPIE & "' and MOUVEMECR = " & xZMOUVEA0.MOUVEMECR
            xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " & xWhere
            Set rsSab = cnsab.Execute(xSQL)
    
            If Not rsSab.EOF Then
                rsYBIAMVT0_GetBuffer rsSab, xYBIAMVT0
                srvYBIAMVT0_fgDisplay xYBIAMVT0, fgSelect_D
                fgSelect_D.Visible = True
            End If
        End If
   End If
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
Me.Enabled = False
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
        Select Case cmdSelect_SQL_K
            Case "1"
                    xZMOUVEA0 = selZMOUVEA0(K)
                    fgDossier_Display
            Case "2"
                    xYBIACPT0 = selYBIACPT0(K)
                    fgDossier_Display_2
            Case "3"
                    xZCHGOPE0 = selZCHGOPE0(K)
                    fgDossier_Display_3
        End Select
   End If
End If
Me.Enabled = True
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

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False
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

Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub

Private Sub mnuContextQuitter_Click()
Unload Me
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim meUnit As typeUnit, X As String
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), SAB_CDRAut)

blnSetfocus = True
Form_Init


blnAuto = False


End Sub


Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
'    cmdlstSourceScan_Click
Else
    If currentAction = "" Then
        If SSTab1.Tab > 0 Then
            SSTab1.Tab = 0
        Else
           'SendKeys "{TAB}"
           ' cmdSelect_Click
        End If
    End If
End If
End Sub









Private Sub mnuPrint0_All_Click()
Dim I As Long, K As Long
Me.Enabled = False: Me.MousePointer = vbHourglass
    
For I = 1 To arrZMOUVEA0_Nb
    fgSelect.Row = I
    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
    xZMOUVEA0 = arrZMOUVEA0(K)
    'prtSAB_CDR_Monitor xZMOUVEA0
Next I

Me.Show

Me.Enabled = True: Me.MousePointer = 0



End Sub




Public Sub cmdPrint_Ok_1()
Dim K As Long, X As String, xSQL As String
Dim wMOUVEMCOM As String
lstSelect.Visible = False
curDB = 0: curCR = 0: wMOUVEMCOM = ""
prtSAB_CPTMVT_Open "Cumul des Mouvements Comptables :Compte / Dossiers"
For K = 1 To selZMOUVEA0_Nb
    If chkSelect = "0" And selZMOUVEA0(K).MOUVEMMON = 0 Then
    Else
        xZMOUVEA0 = selZMOUVEA0(K)
        If xZMOUVEA0.MOUVEMCOM <> wMOUVEMCOM Then
            If wMOUVEMCOM <> "" Then cmdPrint_Ok_1_Total
            curDB = 0: curCR = 0: wMOUVEMCOM = xZMOUVEA0.MOUVEMCOM
        End If
        
        prtSAB_CPTMVT_NewLine
        XPrt.CurrentX = prtMinX + 50: XPrt.Print xZMOUVEA0.MOUVEMCOM;
        xSQL = "select COMPTEINT,COMPTEDEV from " & paramIBM_Library_SAB & ".ZCOMPTE0  where COMPTECOM = '" & xZMOUVEA0.MOUVEMCOM & "'"
        Set rsSab = cnsab.Execute(xSQL)
        
        If Not rsSab.EOF Then XPrt.CurrentX = prtMinX + 2000: XPrt.Print rsSab("COMPTEINT");

        X = Format$(xZMOUVEA0.MOUVEMNUM, "### ### ###")
        XPrt.CurrentX = prtMinX + 6000 - 100 - XPrt.TextWidth(X):
        XPrt.Print X;
        X = Format$(Abs(xZMOUVEA0.MOUVEMMON), "### ### ### ###.00")
        If xZMOUVEA0.MOUVEMMON > 0 Then
            XPrt.CurrentX = prtMinX + 8000 - 100 - XPrt.TextWidth(X)
            curDB = curDB + xZMOUVEA0.MOUVEMMON
        Else
            XPrt.CurrentX = prtMinX + 10000 - 100 - XPrt.TextWidth(X)
            curCR = curCR + xZMOUVEA0.MOUVEMMON
       End If
        XPrt.Print X;
        XPrt.CurrentX = prtMinX + 10000: XPrt.Print dateIBM10(xZMOUVEA0.MOUVEMDTR, True);

    End If
         
Next K
cmdPrint_Ok_1_Total
prtSAB_CPTMVT_Close

lstSelect.Visible = True


End Sub
Public Sub cmdPrint_Ok_2()
Dim K As Long, X As String, xSQL As String
lstSelect.Visible = False
prtSAB_CPTMVT_Open "Etat des dettes et créances ratachées"
For K = 1 To selYBIACPT0_Nb
    If chkSelect = "0" And selYBIACPT0(K).SOLDECEN = 0 Then
    Else
        xYBIACPT0 = selYBIACPT0(K)
        
        prtSAB_CPTMVT_NewLine
        XPrt.CurrentX = prtMinX + 50: XPrt.Print xYBIACPT0.COMPTECOM;
        
        XPrt.CurrentX = prtMinX + 2000: XPrt.Print xYBIACPT0.COMPTEINT;

        X = Format$(Abs(xYBIACPT0.SOLDECEN), "### ### ### ###.00")
        If xYBIACPT0.SOLDECEN > 0 Then
            XPrt.CurrentX = prtMinX + 8000 - 100 - XPrt.TextWidth(X)
        Else
            XPrt.CurrentX = prtMinX + 10000 - 100 - XPrt.TextWidth(X)
        End If
        XPrt.Print X;
        XPrt.CurrentX = prtMinX + 10000: XPrt.Print dateIBM10(xYBIACPT0.SOLDEDMO, True);

    End If
         
Next K
prtSAB_CPTMVT_Close

lstSelect.Visible = True


End Sub

Public Sub lstSelect_All()
Dim blnSelected As Boolean, K As Long
If chkSelect_All = "0" Then
    blnSelected = False
Else
    blnSelected = True
End If
For K = 0 To lstSelect.ListCount - 1
    lstSelect.Selected(K) = blnSelected
Next K
End Sub

Public Sub cmdSelect_SQL_1()
Dim K As Integer, xSQL As String

Do While Not rsSab.EOF
    xZMOUVEA0.MOUVEMNUM = rsSab("MOUVEMNUM")
    xZMOUVEA0.MOUVEMMON = rsSab("MOUVEMMON")
    If selZMOUVEA0(selZMOUVEA0_Nb).MOUVEMNUM <> xZMOUVEA0.MOUVEMNUM Then
         selZMOUVEA0_Nb = selZMOUVEA0_Nb + 1
         If selZMOUVEA0_Nb > selZMOUVEA0_Max Then
             selZMOUVEA0_Max = selZMOUVEA0_Max + 100
             ReDim Preserve selZMOUVEA0(selZMOUVEA0_Max)
         End If
        selZMOUVEA0(selZMOUVEA0_Nb).MOUVEMCOM = rsSab("MOUVEMCOM")
        selZMOUVEA0(selZMOUVEA0_Nb).MOUVEMDTR = rsSab("MOUVEMDTR")
        selZMOUVEA0(selZMOUVEA0_Nb).MOUVEMNUM = xZMOUVEA0.MOUVEMNUM
        selZMOUVEA0(selZMOUVEA0_Nb).MOUVEMMON = xZMOUVEA0.MOUVEMMON
    Else
        selZMOUVEA0(selZMOUVEA0_Nb).MOUVEMMON = selZMOUVEA0(selZMOUVEA0_Nb).MOUVEMMON + xZMOUVEA0.MOUVEMMON
        selZMOUVEA0(selZMOUVEA0_Nb).MOUVEMDTR = rsSab("MOUVEMDTR")
    End If
    rsSab.MoveNext

Loop

For K = 1 To selZMOUVEA0_Nb
    If selZMOUVEA0(K).MOUVEMMON <> 0 Then
        xSQL = "select CDOREGDRE from " & paramIBM_Library_SAB & ".ZCDOREG0 " _
             & "where CDOREGCOP = 'CDE' and CDOREGDOS = " & selZMOUVEA0(K).MOUVEMNUM _
             & " order by CDOREGDRE DESC"
        Set rsSab = cnsab.Execute(xSQL)
        If Not rsSab.EOF Then selZMOUVEA0(K).MOUVEMDTR = rsSab("CDOREGDRE")
    End If
Next K

End Sub

Public Sub cmdSelect_SQL_1_Reprise()
Dim K As Long

Do While Not rsSab.EOF
    xZMOUVEA0.MOUVEMCOM = rsSab("MOUVEMCOM")
    xZMOUVEA0.MOUVEMNUM = rsSab("MOUVEMNUM")
    xZMOUVEA0.MOUVEMMON = rsSab("MOUVEMMON")
    For K = 1 To selZMOUVEA0_Nb
        If selZMOUVEA0(K).MOUVEMNUM = xZMOUVEA0.MOUVEMNUM _
        And selZMOUVEA0(K).MOUVEMCOM = xZMOUVEA0.MOUVEMCOM Then
            selZMOUVEA0(K).MOUVEMMON = selZMOUVEA0(K).MOUVEMMON + xZMOUVEA0.MOUVEMMON
            Exit For
        End If
    Next K
    rsSab.MoveNext

Loop

End Sub


Public Sub cmdSelect_SQL_2_C()
Dim X As String, iLen As Integer
Dim xCOMPTECOM As String
Do While Not rsSab.EOF
    xCOMPTECOM = Trim(rsSab("COMPTECOM"))
    iLen = Len(xCOMPTECOM)
    X = Mid$(xCOMPTECOM, iLen - 2, 3)
    Select Case X
        Case "CAV", "CBO", "DGV", "DOR", "DTT", "DTX", "IDH", "IMP", "LDT", "LDV", "LDX", "LIE", "LOB", "LOR", "NOB", "NOS"

             selYBIACPT0_Nb = selYBIACPT0_Nb + 1
             If selYBIACPT0_Nb > selYBIACPT0_Max Then
                 selYBIACPT0_Max = selYBIACPT0_Max + 100
                 ReDim Preserve selYBIACPT0(selYBIACPT0_Max)
             End If
            selYBIACPT0(selYBIACPT0_Nb).COMPTEINT = rsSab("COMPTEINT")
            selYBIACPT0(selYBIACPT0_Nb).SOLDECEN = rsSab("SOLDECEN") / 1000
            selYBIACPT0(selYBIACPT0_Nb).SOLDEDMO = rsSab("SOLDEDMO")
            selYBIACPT0(selYBIACPT0_Nb).COMPTECOM = xCOMPTECOM
    End Select
    rsSab.MoveNext

Loop

End Sub

Public Sub cmdSelect_SQL_3_C()
Dim V

Do While Not rsSab.EOF
    V = rsZCHGOPE0_GetBuffer(rsSab, xZCHGOPE0)
             selZCHGOPE0_Nb = selZCHGOPE0_Nb + 1
             If selZCHGOPE0_Nb > selZCHGOPE0_Max Then
                 selZCHGOPE0_Max = selZCHGOPE0_Max + 100
                 ReDim Preserve selZCHGOPE0(selZCHGOPE0_Max)
             End If
         selZCHGOPE0(selZCHGOPE0_Nb) = xZCHGOPE0
    rsSab.MoveNext

Loop

End Sub

Public Sub lstSelect_Load_2()
rsZBASTAB0_cboK2 14, lstSelect, ""
Call lst_Scan("CCR", lstSelect)
If lstSelect.ListIndex >= 0 Then lstSelect.Selected(lstSelect.ListIndex) = True
Call lst_Scan("CHB", lstSelect)
If lstSelect.ListIndex >= 0 Then lstSelect.Selected(lstSelect.ListIndex) = True
Call lst_Scan("ICC", lstSelect)
If lstSelect.ListIndex >= 0 Then lstSelect.Selected(lstSelect.ListIndex) = True
Call lst_Scan("INT", lstSelect)
If lstSelect.ListIndex >= 0 Then lstSelect.Selected(lstSelect.ListIndex) = True
lblSelect_Param = "Sélection des comptes se terminant par : " & vbCrLf & "CAV,CBO,DGV,DOR,DTT,DTX,IDH,IMP" & vbCrLf & "LDT,LDV,LDX,LIE,LOB,LOR,NOB,NOS"
lblSelect_Param.Visible = True
lstSelect.Visible = True
cmdSelect_Ok_Caption = "Extraire les comptes"
chkSelect.Visible = True
chkSelect.Caption = "Inclure les comptes soldés"
chkSelect.Value = "1"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
End Sub

Public Sub lstSelect_Load_3()
rsZBASTAB0_cboK2 58, lstSelect, " and (BASTABARG like 'TRF%' or BASTABARG like 'CPT%')"
lstSelect.Visible = True
cmdSelect_Ok_Caption = "Extraire les comptes"
chkSelect.Visible = True
chkSelect.Caption = "Inclure les comptes soldés"
chkSelect.Value = "1"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
chkSelect_All.Caption = "Toutes les opérations du jour"
chkSelect_All.Visible = True
txtSelect_AMJMIN.Visible = True

End Sub

Public Sub cmdPrint_Ok_1_Total()
Dim X As String
prtSAB_CPTMVT_NewLine
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
XPrt.CurrentY = XPrt.CurrentY + 50
X = Format$(Abs(curDB), "### ### ### ###.00")
XPrt.CurrentX = prtMinX + 8000 - 100 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(Abs(curCR), "### ### ### ###.00")
XPrt.CurrentX = prtMinX + 10000 - 100 - XPrt.TextWidth(X)
XPrt.Print X;
prtSAB_CPTMVT_NewLine
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor

End Sub

Public Sub fgSelect_D_ZLIBEL0()
Dim I As Integer
Dim xWhere As String, xSQL As String

xWhere = "Where LIBELPIE = '" & xZMOUVEA0.MOUVEMPIE & "' and LIBELECR = " & xZMOUVEA0.MOUVEMECR
xSQL = "select * from " & paramIBM_Library_SAB & ".ZLIBEL0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)
I = 0
fgSelect_D.Rows = 1
fgSelect_D.Rows = 5
Do While Not rsSab.EOF
    I = I + 1
    fgSelect_D.Row = I
    fgSelect_D.Col = 0: fgSelect_D = rsSab("LIBELECR")
    fgSelect_D.Col = 1: fgSelect_D = rsSab("LIBELNUM")
    fgSelect_D.Col = 2: fgSelect_D = rsSab("LIBELLIB")

    rsSab.MoveNext

Loop
fgSelect_D.Visible = True

End Sub
