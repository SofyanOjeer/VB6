VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBIA_Access 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA Access"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13560
   Icon            =   "BIA_Access.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9270
   ScaleWidth      =   13560
   Begin VB.Frame fraSelect_Options_AMJ 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Période"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7920
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   3615
      Begin MSComCtl2.DTPicker txtSelect_AMJMin 
         Height          =   300
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   12632256
         CustomFormat    =   "dd  MM yyy"
         Format          =   47185923
         CurrentDate     =   38699.44875
         MaxDate         =   401768
         MinDate         =   36526.4425347222
      End
      Begin MSComCtl2.DTPicker txtSelect_AMJMax 
         Height          =   300
         Left            =   2040
         TabIndex        =   13
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   12632256
         CustomFormat    =   "dd  MM yyy"
         Format          =   115736579
         CurrentDate     =   38699.44875
         MaxDate         =   401768
         MinDate         =   36526.4425347222
      End
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   3
      Top             =   0
      Width           =   5175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Sélection"
      TabPicture(0)   =   "BIA_Access.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "........"
      TabPicture(1)   =   "BIA_Access.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13290
         Begin VB.Frame fraSelect 
            Height          =   6855
            Left            =   120
            TabIndex        =   8
            Top             =   1320
            Width           =   13095
            Begin MSFlexGridLib.MSFlexGrid fgSelect 
               Height          =   6465
               Left            =   120
               TabIndex        =   10
               Top             =   240
               Width           =   12840
               _ExtentX        =   22648
               _ExtentY        =   11404
               _Version        =   393216
               Rows            =   1
               Cols            =   13
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
               FormatString    =   "       ||                           |||"
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
         End
         Begin VB.Frame fraSelect_Options 
            Height          =   1092
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   11355
            Begin VB.Frame fraSelect_File 
               Height          =   852
               Left            =   4200
               TabIndex        =   14
               Top             =   120
               Visible         =   0   'False
               Width           =   3372
               Begin VB.TextBox txtSelect_File 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   17
                  Top             =   480
                  Width           =   3012
               End
               Begin VB.ComboBox cboSelect_LIB 
                  Height          =   288
                  Left            =   1200
                  TabIndex        =   16
                  Text            =   "cboSelect_LIB"
                  Top             =   120
                  Width           =   1932
               End
               Begin VB.Label lblSelect_File 
                  Caption         =   "Fichier"
                  Height          =   252
                  Left            =   240
                  TabIndex        =   15
                  Top             =   120
                  Width           =   720
               End
            End
            Begin VB.ComboBox cboSelect_SQL 
               Height          =   315
               Left            =   240
               Sorted          =   -1  'True
               TabIndex        =   9
               Text            =   "cboSelect_SQL"
               Top             =   240
               Width           =   3615
            End
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Exécuter la requête"
            Height          =   645
            Left            =   11880
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   360
            Width           =   1095
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
      Picture         =   "BIA_Access.frx":0044
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.Label libSelect 
      BackColor       =   &H00FFFED9&
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
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4905
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Sql_All 
         Caption         =   "Charger toutes les tables"
      End
      Begin VB.Menu mnuSelect_FICBALP0 
         Caption         =   "extraire C:\Temp\FICBALP0.xls"
      End
      Begin VB.Menu mnux1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
      Begin VB.Menu mnux2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Print_Liste 
         Caption         =   "Imprimer listes "
      End
   End
End
Attribute VB_Name = "frmBIA_Access"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'tvwSelect arborescence
' Node.key  :
'   client      : CLI******
'               : CLI*******ADM         niveau lien
'               : CLI*******ADM-------  niveau client lié
'---------------------------------------------------------

Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim BIA_Access_Aut As typeAuthorization
Dim curX1 As Currency, curX2 As Currency
Dim blnAuto As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim cnAdo As New ADODB.Connection, rsADO As New ADODB.Recordset, errADO As ADODB.Error
Dim blnTransaction As Boolean

Dim cnSAB073Y As New ADODB.Connection, rsSAB073Y As New ADODB.Recordset
'___________________________________________________________________________
Dim appExcel As Excel.Application
Dim wbExcel As Excel.Workbook
Dim wsExcel As Excel.Worksheet
'____________________________________________________________________________

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
        For I = 0 To fgSelect_arrIndex
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.Col = 0
    End If
End If

End Sub
Private Sub fgSelect_Display()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fraSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgselect_Display"
    
'For I = 1 To arrYCLIENA0_NB
         
'    xYCLIENA0 = arrYCLIENA0(I)
'        fgSelect.Rows = fgSelect.Rows + 1
'        fgSelect.Row = fgSelect.Rows - 1
'        fgSelect_DisplayLine I
'Next I

fraSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : "): DoEvents
If fgSelect.Rows > 1 Then
    fgSelect_Sort1 = 0: fgSelect_Sort2 = 3: fgSelect_Sort
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
Dim X As String, lenX As Integer
Dim xSql As String
On Error Resume Next
'fgSelect.Col = 0: fgSelect.Text = dateImp10(xYCLIENA0.CLIENACLI)
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

End Sub


Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect.FormatString = fgSelect_FormatString
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


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), BIA_Access_Aut)
Form_Init


Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case Else: blnAuto = False
End Select


End Sub


Public Sub Form_Init()
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    If Not blnAuto Then MsgBox "paramétrage inconsistant", vbCritical, "frmBIA_Access.paramSAA_Init"
    Unload Me
Else
    lstErr.Clear
End If


blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fraSelect_Options_AMJ.Visible = False
Call DTPicker_Set(txtSelect_AmjMax, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtSelect_AmjMin, mId$(YBIATAB0_DATE_CPT_J, 1, 6) & "01")

fgSelect.Enabled = True

'tableYBase_Open
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

'cmdSelect_Ok_Click


blnControl = True



End Sub


Public Function param_Init()

param_Init = Null
Call lstErr_Clear(lstErr, cmdContext, "BIA_Access : param_init"): DoEvents

fraSelect.Visible = False


cboSelect_SQL.Clear
cboSelect_SQL.AddItem "ZADRESS0"
cboSelect_SQL.AddItem "ZALMMM0"
cboSelect_SQL.AddItem "ZAUTHST0"
cboSelect_SQL.AddItem "ZAUTSYC0"
cboSelect_SQL.AddItem "ZBASFUT0"
cboSelect_SQL.AddItem "ZBALEX10"
cboSelect_SQL.AddItem "ZDWHEHB0"
cboSelect_SQL.AddItem "ZDWHEXP0"
cboSelect_SQL.AddItem "ZDWHOPE0"
cboSelect_SQL.AddItem "ZCDODOS0"
cboSelect_SQL.AddItem "ZCDOUTI0"
cboSelect_SQL.AddItem "ZCGSMM10"
cboSelect_SQL.AddItem "ZCGSMM30"
cboSelect_SQL.AddItem "ZCHGDET0"
cboSelect_SQL.AddItem "ZCLIENA0"
cboSelect_SQL.AddItem "ZCLINEX0"
cboSelect_SQL.AddItem "ZCLIGRP0"
cboSelect_SQL.AddItem "ZCOMPTE0"
cboSelect_SQL.AddItem "ZCOMREF0"
cboSelect_SQL.AddItem "ZDECMOU0"
cboSelect_SQL.AddItem "ZDORCPT0"
cboSelect_SQL.AddItem "ZGAPPIS0"
cboSelect_SQL.AddItem "ZLETCOM0"
cboSelect_SQL.AddItem "ZPLAN0"
cboSelect_SQL.AddItem "ZRELEVE0"
cboSelect_SQL.AddItem "ZSOLDE0"
cboSelect_SQL.AddItem "ZTREOPE0"
cboSelect_SQL.AddItem "ZTITULA0"
cboSelect_SQL.AddItem "ZSWIT001"

cboSelect_SQL.AddItem "YAUTE1I0"
cboSelect_SQL.AddItem "YBIATAB0"
cboSelect_SQL.AddItem "YBIACPT0"
cboSelect_SQL.AddItem "YBIAMON7"
cboSelect_SQL.AddItem "YBIAMVT0"
cboSelect_SQL.AddItem "YBIAMVTH"
cboSelect_SQL.AddItem "YBIASTO0"
cboSelect_SQL.AddItem "YBIARELV"
cboSelect_SQL.AddItem "YCHQMON0"
cboSelect_SQL.AddItem "YGUIMAD0"

cboSelect_SQL.AddItem "FICBALP0"
cboSelect_SQL.AddItem "QRY_ACCFIC"
cboSelect_SQL.AddItem "SCRECHW3"
cboSelect_SQL.AddItem "SCRECHW4"

If BIA_Access_Aut.Xspécial Then
    cboSelect_SQL.AddItem "_Manuel"
    X = MsgBox("inclure les fichiers ZMNU, YROP ?", vbQuestion + vbYesNo, "Extraction SAB073 => .mdb")
    If X = vbYes Then
        cboSelect_SQL.AddItem "ZMNUHLB0"
        cboSelect_SQL.AddItem "ZMNUMEN0"
        cboSelect_SQL.AddItem "ZMNUOPT0"
        cboSelect_SQL.AddItem "ZMNURUT0"
        cboSelect_SQL.AddItem "ZMNUUTI0"
        cboSelect_SQL.AddItem "ZMNUUTP0"
    
        cboSelect_SQL.AddItem "YROPDOS0"
        cboSelect_SQL.AddItem "YROPINF0"
    
        cboSelect_SQL.AddItem "SideEUPLAB0"
        cboSelect_SQL.AddItem "YEUPMON0"
    End If
End If


cboSelect_LIB.Clear
cboSelect_LIB.AddItem "SAB073"
cboSelect_LIB.AddItem "SAB073SPE"
cboSelect_LIB.AddItem "SAB073JRN"
cboSelect_LIB.AddItem "JPLTST"
cboSelect_LIB.ListIndex = 0

'cboSelect_SQL.AddItem "Z0"
cboSelect_SQL.ListIndex = 0

Me.Enabled = True: Me.MousePointer = 0

End Function





Private Sub traite_NEWDATE(tbl As String)
Dim xSql As String

On Error GoTo gere_error

    xSql = "ALTER TABLE " & tbl & " DROP COLUMN [Date d'extraction]"
    rsSAB073Y.Open xSql
    xSql = "ALTER TABLE " & tbl & " ADD COLUMN [Date d'extraction] TEXT(19)"
    rsSAB073Y.Open xSql
    xSql = "update " & tbl & " set [Date d'extraction] = '" & retourne_newDate & "'"
    rsSAB073Y.Open xSql
Exit Sub
gere_error:
    If Err.Number = -2147217900 Then
        Resume Next
    Else
        MsgBox Err.Description
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



Private Sub cboSelect_LIB_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub cboSelect_SQL_Click()
If cboSelect_SQL = "YBIAMVTH" Then fraSelect_Options_AMJ.Visible = True
If cboSelect_SQL = "ZDECMOU0" Then fraSelect_Options_AMJ.Visible = True
If cboSelect_SQL = "ZCHGDET0" Then fraSelect_Options_AMJ.Visible = True

End Sub


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
fraSelect_Options_AMJ.Visible = False
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case SSTab1.Tab
    Case 0:
         'If fgSelect.Rows > 1 Then
             Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
       ' End If
                
    Case 1:
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_SQL()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim wAmj7 As Long
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL"

If fraSelect_File.Visible Then
    Select Case Trim(txtSelect_File)
        Case Else: cmdSelect_TEST_Update
    End Select
    fraSelect_File.Visible = False
    Exit Sub
End If


Select Case cboSelect_SQL.Text
    Case "_Manuel": fraSelect_File.Visible = True
    Case "ZADRESS0": cmdSelect_ZADRESS0
    Case "ZALMMM0": cmdSelect_ZALMMM0
    Case "ZAUTHST0": cmdSelect_ZAUTHST0
    Case "ZAUTSYC0": cmdSelect_ZAUTSYC0
    Case "ZBASFUT0": cmdSelect_ZBASFUT0
    Case "ZBASFUT0": cmdSelect_ZBASFUT0
    Case "ZBALEX10": cmdSelect_ZBALEX10
    Case "ZDWHEHB0": cmdSelect_ZDWHEHB0
    Case "ZDWHEXP0": cmdSelect_ZDWHEXP0
    Case "ZDWHOPE0": cmdSelect_ZDWHOPE0
    Case "ZCDODOS0": cmdSelect_ZCDODOS0
    Case "ZCDOUTI0": cmdSelect_ZCDOUTI0
    Case "ZCGSMM10": cmdSelect_ZCGSMM10
    Case "ZCGSMM30": cmdSelect_ZCGSMM30
    Case "ZCHGDET0": cmdSelect_ZCHGDET0
    Case "ZCLIENA0": cmdSelect_ZCLIENA0
    Case "ZCLINEX0": cmdSelect_ZCLINEX0
    Case "ZCLIGRP0": cmdSelect_ZCLIGRP0
    Case "ZCOMPTE0": cmdSelect_ZCOMPTE0
    Case "ZCOMREF0": cmdSelect_ZCOMREF0
    Case "ZDECMOU0": cmdSelect_ZDECMOU0
    Case "ZDORCPT0": cmdSelect_ZDORCPT0
    Case "ZGAPPIS0": cmdSelect_ZGAPPIS0
    Case "ZLETCOM0": cmdSelect_ZLETCOM0
    Case "ZMNUHLB0": cmdSelect_ZMNUHLB0
    Case "ZMNUMEN0": cmdSelect_ZMNUMEN0
    Case "ZMNUOPT0": cmdSelect_ZMNUOPT0
    Case "ZMNURUT0": cmdSelect_ZMNURUT0
    Case "ZMNUUTI0": cmdSelect_ZMNUUTI0
    Case "ZMNUUTP0": cmdSelect_ZMNUUTP0
    Case "ZPLAN0": cmdSelect_ZPLAN0
    Case "ZRELEVE0": cmdSelect_ZRELEVE0
    Case "ZSOLDE0": cmdSelect_ZSOLDE0: cmdSelect_ZSOLDE0J_1: cmdSelect_ZSOLDE0J_2
    Case "ZTITULA0": cmdSelect_ZTITULA0
    Case "ZTREOPE0": cmdSelect_ZTREOPE0
    Case "ZSWIT001": cmdSelect_ZSWIT001
    
    Case "YAUTE1I0": cmdSelect_YAUTE1I0
    Case "YBIACPT0": cmdSelect_YBIACPT0
    Case "YBIAMON7": cmdSelect_YBIAMON7
    Case "YBIAMVT0": cmdSelect_YBIAMVT0
    Case "YBIAMVTH": cmdSelect_YBIAMVTH
    Case "YBIASTO0": cmdSelect_YBIASTO0
    Case "YBIARELV": cmdSelect_YBIARELV
    Case "YBIATAB0": cmdSelect_YBIATAB0
    Case "YCHQMON0": cmdSelect_YCHQMON0
    Case "YROPDOS0": cmdSelect_YROPDOS0
    Case "YROPINF0": cmdSelect_YROPINF0
    Case "YGUIMAD0": cmdSelect_YGUIMAD0
    Case "ZADRESS0": cmdSelect_ZADRESS0

    Case "FICBALP0": cmdSelect_FICBALP0
    Case "QRY_ACCFIC": cmdSelect_QRY_ACCFIC
    Case "SCRECHW3": cmdSelect_SCRECHW3
    Case "SCRECHW4": cmdSelect_SCRECHW4
    
    Case "SideEUPLAB0": cmdSelect_SideEUPLAB0
    Case "YEUPMON0": cmdSelect_YEUPMON0

End Select
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub
Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_Access_cmdSelect_Ok ........"): DoEvents
fraSelect_Options_AMJ.Visible = False

'fgSelect.Clear
If blnOk Then
    cmdSelect_Ok.Caption = "Options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = False
    cmdSelect_SQL
    MsgBox "Fin de l'extraction..."
Else
    cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = True
    fraSelect.Visible = False
End If
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_Access_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0


End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 5: fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_Sort
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
        Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        'fgSelect.Col = fgSelect_arrIndex:  arrYCLIENA0_Index = CLng(fgSelect.Text)
        fgSelect.LeftCol = 0
       ' oldYCLIENA0 = arrYCLIENA0(arrYCLIENA0_Index)
   End If
End If
fgSelect.LeftCol = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)

cnADO_Close
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
fraSelect_Options_AMJ.Visible = False
lstErr.Clear: lstErr.Height = 200
If SSTab1.Tab = 0 Then
        Unload Me
    Exit Sub
Else
    SSTab1.Tab = SSTab1.Tab - 1
End If

End Sub

Public Sub cmdContext_Return()
fraSelect_Options_AMJ.Visible = False
If SSTab1.Tab = 0 Then
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

cnADO_Open
Exit Sub

Error_Handler:

blnControl = False
If Not blnAuto Then MsgBox Error
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


Private Sub mnuSelect_FICBALP0_Click()
Dim I As Long
Dim xSql As String
Dim xFICBALP0 As typeFICBALP0


Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_Access_mnuSelect_FICBALP0 ........"): DoEvents
Call lstErr_AddItem(lstErr, cmdContext, "> initialisation ......"): DoEvents
'___________________________________________________________________
Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
Set wsExcel = wbExcel.ActiveSheet
'___________________________________________________________________


appExcel.Visible = False


With wsExcel
    .Cells((1), (1)) = "COMPTEDEV"
    .Cells((1), (2)) = "COMPTEOBL"
    .Cells((1), (3)) = "Classe"
    .Cells((1), (4)) = "BIL_HBL"
    .Cells((1), (5)) = "COMPTECOM"
    .Cells((1), (6)) = "COMPTEINT"
    .Cells((1), (7)) = "SOLDE_W"
    .Cells((1), (8)) = "SOLDECVL"
End With

xSql = "select * from " & paramIBM_Library_SABSPE & ".FICBALP0 order by COMPTEDEV, COMPTEOBL, COMPTECOM"
Set rsADO = cnAdo.Execute(xSql)
I = 1

Do While Not rsADO.EOF
    
    Call rsFICBALP0_GetBuffer(rsADO, xFICBALP0)
    I = I + 1
    With wsExcel
        .Cells(I, (1)) = rsADO("COMPTEDEV")
        .Cells(I, (2)) = rsADO("COMPTEOBL")
        .Cells(I, (3)) = rsADO("Classe")
        .Cells(I, (4)) = rsADO("BIL_HBL")
        .Cells(I, (5)) = rsADO("COMPTECOM")
        .Cells(I, (6)) = rsADO("COMPTEINT")
        .Cells(I, (7)) = rsADO("SOLDE_W")
        .Cells(I, (8)) = rsADO("SOLDECVL")
End With
If I Mod 10 Then Call lstErr_ChangeLastItem(lstErr, cmdContext, "> Nb : " & I & " - " & wsExcel.Cells(I, (5))): DoEvents

    rsADO.MoveNext
Loop


MousePointer = 0

appExcel.ActiveWorkbook.SaveAs ("c:\temp\FICBALP0.xls")
appExcel.Visible = True


'___________________________________________________________________
wbExcel.Close
appExcel.Quit

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
'___________________________________________________________________

Call lstErr_AddItem(lstErr, cmdContext, "< BIA_Access_mnuSelect_FICBALP0"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_Liste_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
'cmdPrint_List1_Ok
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Sql_All_Click()
Dim I As Integer
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_Access_cmdSelect_Ok ........"): DoEvents

For I = 1 To cboSelect_SQL.ListCount - 1
    cboSelect_SQL.ListIndex = I
    If cboSelect_SQL <> "YBIAMVTH" And cboSelect_SQL <> "ZCHGDET0" Then
        Call lstErr_AddItem(lstErr, cmdContext, cboSelect_SQL.Text): DoEvents
        cmdSelect_SQL
    End If
Next I
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_Access_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then cmdSelect_Ok.SetFocus

End Sub














Private Sub SSTab1_GotFocus()
Select Case SSTab1.Tab
    Case 0: fgSelect.LeftCol = 0
   ' Case 1: fgSAA.LeftCol = 0
End Select
End Sub


Public Sub blnTransaction_Set()
If Not blnTransaction Then
    blnTransaction = True
   ' Set rsADO_Update = cnado.Execute("SET TRANSACTION ISOLATION LEVEL READ COMMITTED")

End If

End Sub


Public Sub cnADO_Close()
On Error Resume Next

cnAdo.Close
Set cnAdo = Nothing

cnSAB073Y.Close
Set cnSAB073Y = Nothing


End Sub

Public Sub cnADO_Open()
On Error GoTo Error_Handler
Dim X As String

cnAdo.Open paramODBC_DSN_SAB
cnSAB073Y.Open paramODBC_DSN_SAB073Y


Exit Sub

Error_Handler:
blnControl = False
If Not blnAuto Then MsgBox Error

End Sub



Public Sub cmdSelect_ZCLIENA0()
Dim xSql As String
Dim xZCLIENA0 As typeZCLIENA0

xSql = "Delete * from ZCLIENA0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZCLIENA0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZCLIENA0_GetBuffer(rsADO, xZCLIENA0)
    V = adoZCLIENA0_AddNew(rsSAB073Y, xZCLIENA0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZCLIENA0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZCLIENA0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZCHGDET0()
Dim xSql As String
Dim xZCHGDET0 As typeZCHGDET0
Dim wAmjMin As String, wAmjMax As String

xSql = "Delete * from ZCHGDET0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZCHGDET0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

Call DTPicker_Control(txtSelect_AmjMin, wAmjMin)
Call DTPicker_Control(txtSelect_AmjMax, wAmjMax)

    xSql = "select * from " & paramIBM_Library_SAB & ".ZCHGDET0 where CHGDETDTE >= " _
            & wAmjMin - 19000000 & " and CHGDETDTE <= " & wAmjMax - 19000000

'xSql = "select * from " & paramIBM_Library_SAB & ".ZCHGDET0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZCHGDET0_GetBuffer(rsADO, xZCHGDET0)
    V = adoZCHGDET0_AddNew(rsSAB073Y, xZCHGDET0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZCHGDET0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZCHGDET0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZCDODOS0()
Dim xSql As String
Dim xZCDODOS0 As typeZCDODOS0

xSql = "Delete * from ZCDODOS0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZCDODOS0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZCDODOS0_GetBuffer(rsADO, xZCDODOS0)
    V = adoZCDODOS0_AddNew(rsSAB073Y, xZCDODOS0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZCDODOS0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZCDODOS0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZCDOUTI0()
Dim xSql As String
Dim xZCDOUTI0 As typeZCDOUTI0
Dim V
xSql = "Delete * from ZCDOUTI0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZCDOUTI0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZCDOUTI0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZCDOUTI0_GetBuffer(rsADO, xZCDOUTI0)
    V = adoZCDOUTI0_AddNew(rsSAB073Y, xZCDOUTI0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZCDOUTI0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZCDOUTI0")
Set rsSAB073Y = Nothing

End Sub


Public Sub cmdSelect_ZCOMREF0()
Dim xSql As String
Dim xZCOMREF0 As typeZCOMREF0

xSql = "Delete * from ZCOMREF0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZCOMREF0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZCOMREF0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZCOMREF0_GetBuffer(rsADO, xZCOMREF0)
    V = adoZCOMREF0_AddNew(rsSAB073Y, xZCOMREF0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZCOMREF0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZCOMREF0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZRELEVE0()
Dim xSql As String
Dim xZRELEVE0 As typeZRELEVE0

xSql = "Delete * from ZRELEVE0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZRELEVE0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZRELEVE0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZRELEVE0_GetBuffer(rsADO, xZRELEVE0)
    V = adoZRELEVE0_AddNew(rsSAB073Y, xZRELEVE0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZRELEVE0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZRELEVE0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZTREOPE0()
Dim xSql As String
Dim xZTREOPE0 As typeZTREOPE0

xSql = "Delete * from ZTREOPE0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZTREOPE0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZTREOPE0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZTREOPE0_GetBuffer(rsADO, xZTREOPE0)
    V = adoZTREOPE0_AddNew(rsSAB073Y, xZTREOPE0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZTREOPE0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZTREOPE0")
Set rsSAB073Y = Nothing

End Sub
Public Sub cmdSelect_ZBASFUT0()
Dim xSql As String
Dim xZBASFUT0 As typeZBASFUT0

xSql = "Delete * from ZBASFUT0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZBASFUT0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZBASFUT0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZBASFUT0_GetBuffer(rsADO, xZBASFUT0)
    V = adoZBASFUT0_AddNew(rsSAB073Y, xZBASFUT0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZBASFUT0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZBASFUT0")
Set rsSAB073Y = Nothing

End Sub
Public Sub cmdSelect_ZBALEX10()
Dim xSql As String
Dim xZBALEX10 As typeZBALEX10

xSql = "Delete * from ZBALEX10"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZBALEX10", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZBALEX10"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZBALEX10_GetBuffer(rsADO, xZBALEX10)
    V = adoZBALEX10_AddNew(rsSAB073Y, xZBALEX10)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZBALEX10"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZBALEX10")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZDWHEXP0()
Dim xSql As String
Dim xZDWHEXP0 As typeZDWHEXP0

xSql = "Delete * from ZDWHEXP0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZDWHEXP0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZDWHEXP0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZDWHEXP0_GetBuffer(rsADO, xZDWHEXP0)
    V = adoZDWHEXP0_AddNew(rsSAB073Y, xZDWHEXP0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZDWHEXP0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZDWHEXP0")
Set rsSAB073Y = Nothing

End Sub
Public Sub cmdSelect_ZDWHEHB0()
Dim xSql As String
Dim xZDWHEHB0 As typeZDWHEHB0

xSql = "Delete * from ZDWHEHB0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZDWHEHB0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZDWHEHB0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZDWHEHB0_GetBuffer(rsADO, xZDWHEHB0)
    V = adoZDWHEHB0_AddNew(rsSAB073Y, xZDWHEHB0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZDWHEHB0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZDWHEHB0")
Set rsSAB073Y = Nothing

End Sub


Public Sub cmdSelect_ZDWHOPE0()
Dim xSql As String
Dim xZDWHOPE0 As typeZDWHOPE0

xSql = "Delete * from ZDWHOPE0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZDWHOPE0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZDWHOPE0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZDWHOPE0_GetBuffer(rsADO, xZDWHOPE0)
    V = adoZDWHOPE0_AddNew(rsSAB073Y, xZDWHOPE0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZDWHOPE0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZDWHOPE0")
Set rsSAB073Y = Nothing

End Sub


Public Sub cmdSelect_ZAUTSYC0()
Dim xSql As String
Dim xZAUTSYC0 As typeZAUTSYC0

xSql = "Delete * from ZAUTSYC0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZAUTSYC0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZAUTSYC0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZAUTSYC0_GetBuffer(rsADO, xZAUTSYC0)
    V = adoZAUTSYC0_AddNew(rsSAB073Y, xZAUTSYC0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZAUTSYC0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZAUTSYC0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZAUTHST0()
Dim xSql As String
Dim xZAUTHST0 As typeZAUTHST0

xSql = "Delete * from ZAUTHST0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZAUTHST0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZAUTHST0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZAUTHST0_GetBuffer(rsADO, xZAUTHST0)
    V = adoZAUTHST0_AddNew(rsSAB073Y, xZAUTHST0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZAUTHST0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZAUTHST0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_YBIAMVT0()
Dim xSql As String
Dim xYBIAMVT0 As typeYBIAMVT0
'X = MsgBox("yes => SAB073SPE/YBIAMVT0 " & vbCrLf & "no => \\YBASE\...\YBIAMVT0.txt ", vbYesNo, "Chargement du fichier")

If Not blnOff_Line Then
    xSql = "Delete * from YBIAMVT0"
    Call FEU_ROUGE
    Set rsSAB073Y = cnSAB073Y.Execute(xSql)
    Call FEU_VERT
    Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
    Wait_SS 5

End If

rsSAB073Y.Open "select * from YBIAMVT0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic
If blnOff_Line Then
'extrait Khalifa    cmdSelect_HISMVTP0_txt "HISTMVT_25272.csv"

'    cmdSelect_YBIAMVT0_txt "200301_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200302_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200303_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200304_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200305_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200306_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200307_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200308_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200309_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200310_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200311_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200312_YBIAMVT0"

'    cmdSelect_YBIAMVT0_txt "200401_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200402_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200403_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200404_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200405_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200406_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200407_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200408_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200409_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200410_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200411_YBIAMVT0"
'    cmdSelect_YBIAMVT0_txt "200412_YBIAMVT0"
    
    cmdSelect_YBIAMVT0_txt "200501_YBIAMVT0"
    cmdSelect_YBIAMVT0_txt "200502_YBIAMVT0"
    cmdSelect_YBIAMVT0_txt "YBIAMVT0"
Else

    xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTW"
'    xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH where MOUVEMDCO > 1050000 and MOUVEMDCO < 1060000"
    Set rsADO = cnAdo.Execute(xSql)
    Do While Not rsADO.EOF
        
        Call rsYBIAMVT0_GetBuffer(rsADO, xYBIAMVT0)
        V = adoYBIAMVT0_AddNew(rsSAB073Y, xYBIAMVT0)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_YBIAMVT0"
        rsADO.MoveNext
    Loop
End If
rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("YBIAMVT0")
Set rsSAB073Y = Nothing

End Sub
Public Sub cmdSelect_YBIAMVTH()
Dim xSql As String
Dim wAmjMin As String, wAmjMax As String

Dim xYBIAMVT0 As typeYBIAMVT0
xSql = "Delete * from YBIAMVT0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

Call DTPicker_Control(txtSelect_AmjMin, wAmjMin)
Call DTPicker_Control(txtSelect_AmjMax, wAmjMax)

rsSAB073Y.Open "select * from YBIAMVT0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH where MOUVEMDTR >= " _
            & wAmjMin - 19000000 & " and MOUVEMDTR <= " & wAmjMax - 19000000
    'xSql = "select * from JPLTST.YBIAMVT0_x"
    Set rsADO = cnAdo.Execute(xSql)
    Do While Not rsADO.EOF
        
        Call rsYBIAMVT0_GetBuffer(rsADO, xYBIAMVT0)
        V = adoYBIAMVT0_AddNew(rsSAB073Y, xYBIAMVT0)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_YBIAMVT0"
        rsADO.MoveNext
    Loop
rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("YBIAMVT0")
Set rsSAB073Y = Nothing

End Sub


Public Sub cmdSelect_ZDECMOU0()
Dim xSql As String
Dim wAmjMin As String, wAmjMax As String

Dim xZDECMOU0 As typeZDECMOU0
xSql = "Delete * from ZDECMOU0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

Call DTPicker_Control(txtSelect_AmjMin, wAmjMin)
Call DTPicker_Control(txtSelect_AmjMax, wAmjMax)

rsSAB073Y.Open "select * from ZDECMOU0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic
    xSql = "select * from " & paramIBM_Library_SAB & ".ZDECMOU0 where DECMOUDTR >= " _
            & wAmjMin - 19000000 & " and DECMOUDTR <= " & wAmjMax - 19000000
    Set rsADO = cnAdo.Execute(xSql)
    Do While Not rsADO.EOF
        
        Call rsZDECMOU0_GetBuffer(rsADO, xZDECMOU0)
        V = adoZDECMOU0_AddNew(rsSAB073Y, xZDECMOU0)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZDECMOU0"
        rsADO.MoveNext
    Loop
rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZDECMOU0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZSOLDE0()
Dim xSql As String
Dim xZSOLDE0 As typeZSOLDE0

xSql = "Delete * from ZSOLDE0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZSOLDE0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZSOLDE0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZSOLDE0_GetBuffer(rsADO, xZSOLDE0)
    V = adoZSOLDE0_AddNew(rsSAB073Y, xZSOLDE0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZSOLDE0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZSOLDE0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZSOLDE0J_1()
Dim xSql As String
Dim xZSOLDE0 As typeZSOLDE0

xSql = "Delete * from ZSOLDE0J_1"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZSOLDE0J_1", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SABSPE & ".ZSOLDE0J_1"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZSOLDE0_GetBuffer(rsADO, xZSOLDE0)
    V = adoZSOLDE0_AddNew(rsSAB073Y, xZSOLDE0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZSOLDE0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZSOLDE0J_1")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZSOLDE0J_2()
Dim xSql As String
Dim xZSOLDE0 As typeZSOLDE0

xSql = "Delete * from ZSOLDE0J_2"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZSOLDE0J_2", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SABSPE & ".ZSOLDE0J_2"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZSOLDE0_GetBuffer(rsADO, xZSOLDE0)
    V = adoZSOLDE0_AddNew(rsSAB073Y, xZSOLDE0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZSOLDE0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZSOLDE0J_2")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZTITULA0()
Dim xSql As String
Dim xZTITULA0 As typeZTITULA0

xSql = "Delete * from ZTITULA0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZTITULA0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZTITULA0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZTITULA0_GetBuffer(rsADO, xZTITULA0)
    V = adoZTITULA0_AddNew(rsSAB073Y, xZTITULA0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZTITULA0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZTITULA0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZPLAN0()
Dim xSql As String
Dim xZPLAN0 As typeZPLAN0

xSql = "Delete * from ZPLAN0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZPLAN0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZPLAN0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZPLAN0_GetBuffer(rsADO, xZPLAN0)
    V = adoZPLAN0_AddNew(rsSAB073Y, xZPLAN0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZPLAN0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZPLAN0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZDORCPT0()
Dim xSql As String
Dim xZDORCPT0 As typeZDORCPT0

xSql = "Delete * from ZDORCPT0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZDORCPT0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZDORCPT0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZDORCPT0_GetBuffer(rsADO, xZDORCPT0)
    V = adoZDORCPT0_AddNew(rsSAB073Y, xZDORCPT0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZDORCPT0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZDORCPT0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZLETCOM0()
Dim xSql As String
Dim xZLETCOM0 As typeZLETCOM0

xSql = "Delete * from ZLETCOM0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZLETCOM0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZLETCOM0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZLETCOM0_GetBuffer(rsADO, xZLETCOM0)
    V = adoZLETCOM0_AddNew(rsSAB073Y, xZLETCOM0)
    If Not IsNull(V) Then
        MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZLETCOM0"
        Exit Do
    End If
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZLETCOM0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZGAPPIS0()
Dim xSql As String
Dim xZGAPPIS0 As typeZGAPPIS0

xSql = "Delete * from ZGAPPIS0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZGAPPIS0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZGAPPIS0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZGAPPIS0_GetBuffer(rsADO, xZGAPPIS0)
    V = adoZGAPPIS0_AddNew(rsSAB073Y, xZGAPPIS0)
    If Not IsNull(V) Then
        MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZGAPPIS0"
        Exit Do
    End If
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZGAPPIS0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZMNUOPT0()
Dim xSql As String
Dim xZMNUOPT0 As typeZMNUOPT0

xSql = "Delete * from ZMNUOPT0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZMNUOPT0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic
'X = MsgBox("yes => SAB073SPE/ZMNUOPT0 " & vbCrLf & "no => \\YBASE\...\ZMNUOPT0.txt ", vbYesNo, "Chargement du fichier")
If blnOff_Line Then
    cmdSelect_ZMNUOPT0_txt
Else

    xSql = "select * from " & paramIBM_Library_SAB & ".ZMNUOPT0"
    Set rsADO = cnAdo.Execute(xSql)
    Do While Not rsADO.EOF
        
        Call rsZMNUOPT0_GetBuffer(rsADO, xZMNUOPT0)
        V = adoZMNUOPT0_AddNew(rsSAB073Y, xZMNUOPT0)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZMNUOPT0"
        rsADO.MoveNext
    Loop
End If

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZMNUOPT0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZMNURUT0()
Dim xSql As String
Dim xZMNURUT0 As typeZMNURUT0

xSql = "Delete * from ZMNURUT0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZMNURUT0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZMNURUT0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZMNURUT0_GetBuffer(rsADO, xZMNURUT0)
    V = adoZMNURUT0_AddNew(rsSAB073Y, xZMNURUT0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZMNURUT0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZMNURUT0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZMNUUTI0()
Dim xSql As String
Dim xZMNUUTI0 As typeZMNUUTI0

xSql = "Delete * from ZMNUUTI0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5
rsSAB073Y.Open "select * from ZMNUUTI0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic
'X = MsgBox("yes => SAB073SPE/ZMNUUTI0 " & vbCrLf & "no => \\YBASE\...\ZMNUUTI0.txt ", vbYesNo, "Chargement du fichier")
If blnOff_Line Then
    cmdSelect_ZMNUUTI0_txt
Else

    xSql = "select * from " & paramIBM_Library_SAB & ".ZMNUUTI0"
    Set rsADO = cnAdo.Execute(xSql)
    Do While Not rsADO.EOF
        
        Call rsZMNUUTI0_GetBuffer(rsADO, xZMNUUTI0)
        V = adoZMNUUTI0_AddNew(rsSAB073Y, xZMNUUTI0)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZMNUUTI0"
        rsADO.MoveNext
    Loop
End If
rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZMNUUTI0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZMNUMEN0()
Dim xSql As String
Dim xZMNUMEN0 As typeZMNUMEN0

xSql = "Delete * from ZMNUMEN0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZMNUMEN0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic
'X = MsgBox("yes => SAB073SPE/ZMNUMEN0 " & vbCrLf & "no => \\YBASE\...\ZMNUMEN0.txt ", vbYesNo, "Chargement du fichier")
If blnOff_Line Then
    cmdSelect_ZMNUMEN0_txt
Else

    xSql = "select * from " & paramIBM_Library_SAB & ".ZMNUMEN0"
    Set rsADO = cnAdo.Execute(xSql)
    Do While Not rsADO.EOF
        
        Call rsZMNUMEN0_GetBuffer(rsADO, xZMNUMEN0)
        V = adoZMNUMEN0_AddNew(rsSAB073Y, xZMNUMEN0)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZMNUMEN0"
        rsADO.MoveNext
    Loop
End If
rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZMNUMEN0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZMNUUTP0()
Dim xSql As String
Dim xZMNUUTP0 As typeZMNUUTP0

xSql = "Delete * from ZMNUUTP0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZMNUUTP0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic
'X = MsgBox("yes => SAB073SPE/ZMNUUTP0 " & vbCrLf & "no => \\YBASE\...\ZMNUUTP0.txt ", vbYesNo, "Chargement du fichier")
If blnOff_Line Then
    cmdSelect_ZMNUUTP0_txt
Else

    xSql = "select * from " & paramIBM_Library_SAB & ".ZMNUUTP0"
    Set rsADO = cnAdo.Execute(xSql)
    Do While Not rsADO.EOF
        
        Call rsZMNUUTP0_GetBuffer(rsADO, xZMNUUTP0)
        V = adoZMNUUTP0_AddNew(rsSAB073Y, xZMNUUTP0)
        If Not IsNull(V) Then
            MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZMNUUTP0"
            Exit Do
        End If
        rsADO.MoveNext
    Loop
End If
rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZMNUUTP0")
Set rsSAB073Y = Nothing

End Sub


Public Sub cmdSelect_ZMNUHLB0()
Dim xSql As String
Dim xZMNUHLB0 As typeZMNUHLB0

xSql = "Delete * from ZMNUHLB0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZMNUHLB0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic
'X = MsgBox("yes => SAB073SPE/ZMNUHLB0 " & vbCrLf & "no => \\YBASE\...\ZMNUHLB0.txt ", vbYesNo, "Chargement du fichier")
If blnOff_Line Then
    cmdSelect_ZMNUHLB0_txt
Else

    xSql = "select * from " & paramIBM_Library_SAB & ".ZMNUHLB0"
    Set rsADO = cnAdo.Execute(xSql)
    Do While Not rsADO.EOF
        
        Call rsZMNUHLB0_GetBuffer(rsADO, xZMNUHLB0)
        V = adoZMNUHLB0_AddNew(rsSAB073Y, xZMNUHLB0)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZMNUHLB0"
        rsADO.MoveNext
    Loop
End If
rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZMNUHLB0")
Set rsSAB073Y = Nothing

End Sub


Public Sub cmdSelect_ZCOMPTE0()
Dim xSql As String
Dim xZCOMPTE0 As typeZCOMPTE0

xSql = "Delete * from ZCOMPTE0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZCOMPTE0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZCOMPTE0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZCOMPTE0_GetBuffer(rsADO, xZCOMPTE0)
    V = adoZCOMPTE0_AddNew(rsSAB073Y, xZCOMPTE0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZCOMPTE0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZCOMPTE0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZCLIGRP0()
Dim xSql As String
Dim xZCLIGRP0 As typeZCLIGRP0

xSql = "Delete * from ZCLIGRP0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZCLIGRP0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZCLIGRP0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZCLIGRP0_GetBuffer(rsADO, xZCLIGRP0)
    V = adoZCLIGRP0_AddNew(rsSAB073Y, xZCLIGRP0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZCLIGRP0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZCLIGRP0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZCLINEX0()
Dim xSql As String
Dim xZCLINEX0 As typeZCLINEX0

xSql = "Delete * from ZCLINEX0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZCLINEX0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZCLINEX0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZCLINEX0_GetBuffer(rsADO, xZCLINEX0)
    V = adoZCLINEX0_AddNew(rsSAB073Y, xZCLINEX0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZCLINEX0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZCLINEX0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZADRESS0()
Dim xSql As String
Dim xZADRESS0 As typeZADRESS0

xSql = "Delete * from ZADRESS0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZADRESS0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZADRESS0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZADRESS0_GetBuffer(rsADO, xZADRESS0)
    V = adoZADRESS0_AddNew(rsSAB073Y, xZADRESS0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZADRESS0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZADRESS0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_ZALMMM0()
Dim xSql As String
Dim xZALMMM0 As typeZALMMM0

xSql = "Delete * from ZALMMM0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZALMMM0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZALMMM0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsZALMMM0_GetBuffer(rsADO, xZALMMM0)
    V = adoZALMMM0_AddNew(rsSAB073Y, xZALMMM0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZALMMM0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZALMMM0")
Set rsSAB073Y = Nothing

End Sub
Public Sub cmdSelect_SCRECHW4()
Dim xSql As String
Dim xSCRECHW4 As typeSCRECHW4

xSql = "Delete * from SCRECHW4"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from SCRECHW4", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SABSPE & ".SCRECHW4"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsSCRECHW4_GetBuffer(rsADO, xSCRECHW4)
    V = adoSCRECHW4_AddNew(rsSAB073Y, xSCRECHW4)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_SCRECHW4"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("SCRECHW4")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_SCRECHW3()
Dim xSql As String
Dim xSCRECHW3 As typeSCRECHW3

xSql = "Delete * from SCRECHW3"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from SCRECHW3", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SABSPE & ".SCRECHW3"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsSCRECHW3_GetBuffer(rsADO, xSCRECHW3)
    V = adoSCRECHW3_AddNew(rsSAB073Y, xSCRECHW3)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_SCRECHW3"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("SCRECHW3")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_QRY_ACCFIC()
Dim xSql As String
Dim xQRY_ACCFIC As typeQRY_ACCFIC

xSql = "Delete * from QRY_ACCFIC"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from QRY_ACCFIC", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SABSPE & ".QRY_ACCFIC"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsQRY_ACCFIC_GetBuffer(rsADO, xQRY_ACCFIC)
    V = adoQRY_ACCFIC_AddNew(rsSAB073Y, xQRY_ACCFIC)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_QRY_ACCFIC"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("QRY_ACCFIC")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_FICBALP0()
Dim xSql As String
Dim xFICBALP0 As typeFICBALP0

xSql = "Delete * from FICBALP0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from FICBALP0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SABSPE & ".FICBALP0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsFICBALP0_GetBuffer(rsADO, xFICBALP0)
    V = adoFICBALP0_AddNew(rsSAB073Y, xFICBALP0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_FICBALP0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("FICBALP0")
Set rsSAB073Y = Nothing

End Sub

Public Sub cmdSelect_YBIACPT0()
Dim xSql As String
Dim xYBIACPT0 As typeYBIACPT0

xSql = "Delete * from YBIACPT0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from YBIACPT0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsYBIACPT0_GetBuffer(rsADO, xYBIACPT0)
    V = adoYBIACPT0_AddNew(rsSAB073Y, xYBIACPT0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_YBIACPT0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("YBIACPT0")
Set rsSAB073Y = Nothing
End Sub

Public Sub cmdSelect_YAUTE1I0()
Dim xSql As String
Dim xYAUTE1I0 As typeYAUTE1I0

xSql = "Delete * from YAUTE1I0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from YAUTE1I0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SABSPE & ".YAUTE1I0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsYAUTE1I0_GetBuffer(rsADO, xYAUTE1I0)
    V = adoYAUTE1I0_AddNew(rsSAB073Y, xYAUTE1I0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_YAUTE1I0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("YAUTE1I0")
Set rsSAB073Y = Nothing
End Sub

Public Sub cmdSelect_YBIATAB0()
Dim xSql As String
Dim xYBIATAB0 As typeYBIATAB0
Dim X As String

xSql = "Delete * from YBIATAB0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from YBIATAB0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic
'X = MsgBox("yes => SAB073SPE/YBIATAB0 " & vbCrLf & "no => \\YBASE\...\ybiatab0.txt ", vbYesNo, "Chargement du fichier")
If blnOff_Line Then
    cmdSelect_YBIATAB0_txt
Else
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0"
    Set rsADO = cnAdo.Execute(xSql)
    Do While Not rsADO.EOF
        
        Call rsYBIATAB0_GetBuffer(rsADO, xYBIATAB0)
        V = adoYBIATAB0_AddNew(rsSAB073Y, xYBIATAB0)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_YBIATAB0"
        rsADO.MoveNext
    Loop
End If
rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("YBIATAB0")
Set rsSAB073Y = Nothing
End Sub
Public Sub cmdSelect_TEST_Update()
Dim wFile As String
Dim xSql As String
Dim X As String
Dim arrWHFLDI() As String, Nb As Integer, I As Integer, NbLu As Long
'====================================================================
On Error GoTo Error_Handler

wFile = UCase$(Trim(txtSelect_File))
xSql = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".DSPFFDW0K " _
    & " where WHSYSN = '" & paramIBM_AS400_ID & "'" _
    & " and WHLIB = '" & UCase$(cboSelect_LIB) & "'" _
    & " and WHFILE = '" & wFile & "'"
    
Set rsADO = cnAdo.Execute(xSql)
Nb = rsADO("Tally")
ReDim arrWHFLDI(Nb + 10)
xSql = "select * from " & paramIBM_Library_SABSPE & ".DSPFFDW0K " _
    & " where WHSYSN = '" & paramIBM_AS400_ID & "'" _
    & " and WHLIB = '" & UCase$(cboSelect_LIB) & "'" _
    & " and WHFILE = '" & wFile & "'" _
    & " order by WHFOBO"

Set rsADO = cnAdo.Execute(xSql)
Nb = 0
Do While Not rsADO.EOF
    Nb = Nb + 1
    arrWHFLDI(Nb) = Trim(rsADO("WHFLDI"))
    rsADO.MoveNext
Loop



'====================================================================
xSql = "Delete * from " & wFile
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

Call lstErr_AddItem(lstErr, cmdContext, "> début : ")

NbLu = 0
rsSAB073Y.Open "select * from " & wFile, paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic
xSql = "select * from " & UCase$(cboSelect_LIB) & "." & wFile
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    NbLu = NbLu + 1
    If NbLu Mod 1000 = 0 Then Call lstErr_ChangeLastItem(lstErr, cmdContext, "> Lecture : " & NbLu)

    rsSAB073Y.AddNew
    For I = 1 To Nb
        X = arrWHFLDI(I)
        On Error Resume Next
        rsSAB073Y(X) = rsADO(X)
        If Error <> "" Then
            MsgBox Error, vbCritical, "Lu :" & NbLu & " champ : " & I & " " & X
        End If
    Next I
    rsSAB073Y.Update

    rsADO.MoveNext
Loop
Call lstErr_AddItem(lstErr, cmdContext, "< Terminé : " & NbLu)

'On Error GoTo Error_Handler

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE(wFile)
Set rsSAB073Y = Nothing
Exit Sub

Error_Handler:

MsgBox Error, vbCritical, "frmBIA_Access"
End Sub


Public Sub cmdSelect_ZCGSMM30()
Dim xSql As String
Dim xZCGSMM30 As typeZCGSMM30
Dim X As String

xSql = "Delete * from ZCGSMM30"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZCGSMM30", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic
'X = MsgBox("yes => SAB073SPE/ZCGSMM30 " & vbCrLf & "no => \\YBASE\...\ZCGSMM30.txt ", vbYesNo, "Chargement du fichier")
    xSql = "select * from " & paramIBM_Library_SAB & ".ZCGSMM30"
    Set rsADO = cnAdo.Execute(xSql)
    Do While Not rsADO.EOF
        
        Call rsZCGSMM30_GetBuffer(rsADO, xZCGSMM30)
        V = adoZCGSMM30_AddNew(rsSAB073Y, xZCGSMM30)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZCGSMM30"
        rsADO.MoveNext
    Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZCGSMM30")
Set rsSAB073Y = Nothing
End Sub

Public Sub cmdSelect_ZCGSMM10()
Dim xSql As String
Dim xZCGSMM10 As typeZCGSMM10
Dim X As String

xSql = "Delete * from ZCGSMM10"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZCGSMM10", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic
'X = MsgBox("yes => SAB073SPE/ZCGSMM10 " & vbCrLf & "no => \\YBASE\...\ZCGSMM10.txt ", vbYesNo, "Chargement du fichier")
    xSql = "select * from " & paramIBM_Library_SAB & ".ZCGSMM10"
    Set rsADO = cnAdo.Execute(xSql)
    Do While Not rsADO.EOF
        
        Call rsZCGSMM10_GetBuffer(rsADO, xZCGSMM10)
        V = adoZCGSMM10_AddNew(rsSAB073Y, xZCGSMM10)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZCGSMM10"
        rsADO.MoveNext
    Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZCGSMM10")
Set rsSAB073Y = Nothing
End Sub

Public Sub cmdSelect_YBIATAB0_txt()
Dim V, xIn As String, intFile As Integer
Dim xYBIATAB0 As typeYBIATAB0

xIn = paramYBase_DataF & "YBIATAB0" & paramYBase_Data_ExtensionP
intFile = FreeFile(0)
Open xIn For Input As #intFile
Do Until EOF(1)
    Line Input #intFile, xIn
    If Trim(xIn) <> "" Then
        xYBIATAB0.BIATABID = mId$(xIn, 1, 12)
        xYBIATAB0.BIATABK1 = mId$(xIn, 13, 12)
        xYBIATAB0.BIATABK2 = mId$(xIn, 25, 12)
        xYBIATAB0.BIATABTXT = mId$(xIn, 37, 128)
        V = adoYBIATAB0_AddNew(rsSAB073Y, xYBIATAB0)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_YBIATAB0"

    End If
    
Loop
Close intFile

End Sub

Public Sub cmdSelect_YBIAMVT0_txt(lFile As String)
Dim V, xIn As String, intFile As Integer
Dim xYBIAMVT0 As typeYBIAMVT0

xIn = paramYBase_DataF & lFile & paramYBase_Data_ExtensionP
intFile = FreeFile(0)
Open xIn For Input As #intFile
Do Until EOF(1)
    Line Input #intFile, xIn
    If Trim(xIn) <> "" Then
   
            xYBIAMVT0.MOUVEMETA = CInt(Val(mId$(xIn, 1, 5)))
            xYBIAMVT0.MOUVEMPLA = CLng(Val(mId$(xIn, 6, 4)))
            xYBIAMVT0.MOUVEMCOM = mId$(xIn, 10, 20)
            xYBIAMVT0.MOUVEMMON = CCur(mId$(xIn, 30, 18)) / 1000
            xYBIAMVT0.MOUVEMDOP = CLng(Val(mId$(xIn, 48, 8)))
            xYBIAMVT0.MOUVEMDVA = CLng(Val(mId$(xIn, 56, 8)))
            xYBIAMVT0.MOUVEMDCO = CLng(Val(mId$(xIn, 64, 8)))
            xYBIAMVT0.MOUVEMDTR = CLng(Val(mId$(xIn, 72, 8)))
            xYBIAMVT0.MOUVEMPIE = CLng(Val(mId$(xIn, 80, 10)))
            xYBIAMVT0.MOUVEMECR = CLng(Val(mId$(xIn, 90, 8)))
            xYBIAMVT0.MOUVEMOPE = mId$(xIn, 98, 3)
            xYBIAMVT0.MOUVEMNUM = CLng(Val(mId$(xIn, 101, 10)))
            xYBIAMVT0.MOUVEMSCH = CInt(Val(mId$(xIn, 111, 5)))
            xYBIAMVT0.MOUVEMUTI = CInt(Val(mId$(xIn, 116, 5)))
            xYBIAMVT0.MOUVEMAGE = CInt(Val(mId$(xIn, 121, 5)))
            xYBIAMVT0.MOUVEMSER = mId$(xIn, 126, 2)
            xYBIAMVT0.MOUVEMSSE = mId$(xIn, 128, 2)
            xYBIAMVT0.MOUVEMEXO = mId$(xIn, 130, 1)
            xYBIAMVT0.MOUVEMANA = mId$(xIn, 131, 6)
            xYBIAMVT0.MOUVEMBDF = mId$(xIn, 137, 3)
            xYBIAMVT0.MOUVEMANU = mId$(xIn, 140, 1)
            xYBIAMVT0.MOUVEMRET = mId$(xIn, 141, 1)
            xYBIAMVT0.MOUVEMEVE = mId$(xIn, 142, 3)
            xYBIAMVT0.MOUVEMSAN = mId$(xIn, 145, 6)
            xYBIAMVT0.MOUVEMSAD = mId$(xIn, 151, 80)
            
            xYBIAMVT0.LIBELLIB1 = mId$(xIn, 231, 30)
            xYBIAMVT0.LIBELLIB2 = mId$(xIn, 261, 30)
            xYBIAMVT0.LIBELLIB3 = mId$(xIn, 291, 30)
            xYBIAMVT0.LIBELLIB4 = mId$(xIn, 321, 30)
                
            xYBIAMVT0.COMPTEOBL = mId$(xIn, 351, 10)
            xYBIAMVT0.COMPTEINT = mId$(xIn, 361, 32)
            xYBIAMVT0.COMPTEDEV = mId$(xIn, 393, 3)
            xYBIAMVT0.COMPTELOR = mId$(xIn, 396, 1)
            xYBIAMVT0.COMPTECLA = CLng(Val(mId$(xIn, 397, 3)))
            If IsNumeric(mId$(xIn, 400, 19)) Then
                xYBIAMVT0.BIAMVTSD0 = CCur(mId$(xIn, 400, 19)) / 100
                xYBIAMVT0.BIAMVTID = CLng(Val(mId$(xIn, 419, 11)))
            Else
                xYBIAMVT0.BIAMVTSD0 = 0
                xYBIAMVT0.BIAMVTID = 0
                MsgBox "YBIAMVT0 ???? " & xIn, vbCritical
            End If

        V = adoYBIAMVT0_AddNew(rsSAB073Y, xYBIAMVT0)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_YBIAMVT0"

    End If
    
Loop
Close intFile

End Sub

Public Sub cmdSelect_ZMNUUTI0_txt()
Dim V, xIn As String, intFile As Integer
Dim xZMNUUTI0 As typeZMNUUTI0
Dim K As Integer
xIn = paramYBase_DataF & "ZMNUUTI0.CSV"
intFile = FreeFile(0)
Open xIn For Input As #intFile
Do Until EOF(1)
    Line Input #intFile, xIn
    If Trim(xIn) <> "" Then
        K = 0
            xZMNUUTI0.MNUUTIETB = CInt(Val(CSV_Scan(xIn, K)))
            xZMNUUTI0.MNUUTIREF = CLng(Val(CSV_Scan(xIn, K)))
            xZMNUUTI0.MNUUTICUT = CInt(Val(CSV_Scan(xIn, K)))
            xZMNUUTI0.MNUUTIGR2 = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUUTI0.MNUUTIGR3 = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUUTI0.MNUUTIGR4 = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUUTI0.MNUUTIOUT = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUUTI0.MNUUTILAN = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUUTI0.MNUUTIMSE = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUUTI0.MNUUTIAGE = CInt(Val(CSV_Scan(xIn, K)))
            xZMNUUTI0.MNUUTISER = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUUTI0.MNUUTISRV = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUUTI0.MNUUTIGRS = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUUTI0.MNUUTIGEN = CInt(Val(CSV_Scan(xIn, K)))
            xZMNUUTI0.MNUUTIPOS = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUUTI0.MNUUTIMAI = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))

           ' xZMNUUTI0.MNUUTIETB = CInt(Val(mId$(xIn, 1, 6)))
           ' xZMNUUTI0.MNUUTICUT = CLng(Val(mId$(xIn, 7, 6)))
           ' xZMNUUTI0.MNUUTIGR2 = CLng(Val(mId$(xIn, 13, 6)))
           ' xZMNUUTI0.MNUUTIDRG = mId$(xIn, 19, 1)
           ' xZMNUUTI0.MNUUTIOUT = mId$(xIn, 20, 10)
           ' xZMNUUTI0.MNUUTILAN = mId$(xIn, 30, 1)
           ' xZMNUUTI0.MNUUTIMSE = mId$(xIn, 31, 1)
           ' xZMNUUTI0.MNUUTIAGE = CLng(Val(mId$(xIn, 32, 6)))
           ' xZMNUUTI0.MNUUTISER = mId$(xIn, 38, 2)
           ' xZMNUUTI0.MNUUTISRV = mId$(xIn, 40, 2)

        V = adoZMNUUTI0_AddNew(rsSAB073Y, xZMNUUTI0)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZMNUUTI0"

    End If
    
Loop
Close intFile

End Sub
Public Sub cmdSelect_ZMNUMEN0_txt()
Dim V, xIn As String, intFile As Integer
Dim xZMNUMEN0 As typeZMNUMEN0
Dim K As Integer

xIn = paramYBase_DataF & "ZMNUMEN0.csv"
intFile = FreeFile(0)
Open xIn For Input As #intFile
Do Until EOF(1)
    Line Input #intFile, xIn
    If Trim(xIn) <> "" Then
               K = 0
            xZMNUMEN0.MNUMENETB = CInt(Val(CSV_Scan(xIn, K)))
            xZMNUMEN0.MNUMENREF = CLng(Val(CSV_Scan(xIn, K)))
            xZMNUMEN0.MNUMENGRP = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUMEN0.MNUMENPRE = CLng(Val(CSV_Scan(xIn, K)))
            xZMNUMEN0.MNUMENORD = CLng(Val(CSV_Scan(xIn, K)))
            xZMNUMEN0.MNUMENCOD = CLng(Val(CSV_Scan(xIn, K)))
            xZMNUMEN0.MNUMENOIA = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUMEN0.MNUMENJOQ = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))

            'xZMNUMEN0.MNUMENETB = CInt(Val(mId$(xIn, 1, 6)))
            'xZMNUMEN0.MNUMENCGR = CLng(Val(mId$(xIn, 7, 6)))
            'xZMNUMEN0.MNUMENPRE = CLng(Val(mId$(xIn, 13, 8)))
            'xZMNUMEN0.MNUMENORD = CLng(Val(mId$(xIn, 21, 6)))
            'xZMNUMEN0.MNUMENCOD = CLng(Val(mId$(xIn, 27, 8)))
            'xZMNUMEN0.MNUMENOIA = mId$(xIn, 35, 1)
            'xZMNUMEN0.MNUMENJOQ = mId$(xIn, 36, 10)

        V = adoZMNUMEN0_AddNew(rsSAB073Y, xZMNUMEN0)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZMNUMEN0"

    End If
    
Loop
Close intFile

End Sub
Public Sub cmdSelect_ZMNUUTP0_txt()
Dim V, xIn As String, intFile As Integer
Dim xZMNUUTP0 As typeZMNUUTP0
Dim K As Integer

xIn = paramYBase_DataF & "ZMNUUTP0.csv"
intFile = FreeFile(0)
Open xIn For Input As #intFile
Do Until EOF(1)
    Line Input #intFile, xIn
    If Trim(xIn) <> "" Then
               K = 0
            xZMNUUTP0.MNUUTPETB = CInt(Val(CSV_Scan(xIn, K)))
            xZMNUUTP0.MNUUTPREF = CLng(Val(CSV_Scan(xIn, K)))
            xZMNUUTP0.MNUUTPGRP = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUUTP0.MNUUTPAGE = CInt(Val(CSV_Scan(xIn, K)))
            xZMNUUTP0.MNUUTPOIA = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUUTP0.MNUUTPCLA = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))

        V = adoZMNUUTP0_AddNew(rsSAB073Y, xZMNUUTP0)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZMNUUTP0"

    End If
    
Loop
Close intFile

End Sub

Public Sub cmdSelect_ZMNUHLB0_txt()
Dim V, xIn As String, intFile As Integer
Dim xZMNUHLB0 As typeZMNUHLB0
Dim K As Integer

xIn = paramYBase_DataF & "ZMNUHLB0.csv"
intFile = FreeFile(0)
Open xIn For Input As #intFile
Do Until EOF(1)
    Line Input #intFile, xIn
    If Trim(xIn) <> "" Then
               K = 0
            xZMNUHLB0.MNUHLBETB = CInt(Val(CSV_Scan(xIn, K)))
            xZMNUHLB0.MNUHLBREF = CLng(Val(CSV_Scan(xIn, K)))
            xZMNUHLB0.MNUHLBCLA = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUHLB0.MNUHLBNOM = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUHLB0.MNUHLBVAL = CSV_Scan(xIn, K)     '''= EBCDIC_ASCII(CSV_Scan(xIn, K))
            xZMNUHLB0.MNUHLBDBD = CLng(Val(CSV_Scan(xIn, K)))
            xZMNUHLB0.MNUHLBDBH = CLng(Val(CSV_Scan(xIn, K)))
            xZMNUHLB0.MNUHLBFID = CLng(Val(CSV_Scan(xIn, K)))
            xZMNUHLB0.MNUHLBFIH = CLng(Val(CSV_Scan(xIn, K)))
            xZMNUHLB0.MNUHLBSUS = CInt(Val(CSV_Scan(xIn, K)))
            xZMNUHLB0.MNUHLBSDT = CLng(Val(CSV_Scan(xIn, K)))
            xZMNUHLB0.MNUHLBSHE = CLng(Val(CSV_Scan(xIn, K)))


        V = adoZMNUHLB0_AddNew(rsSAB073Y, xZMNUHLB0)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZMNUHLB0"

    End If
    
Loop
Close intFile

End Sub

Public Sub cmdSelect_ZMNUOPT0_txt()
Dim V, xIn As String, intFile As Integer
Dim xZMNUOPT0 As typeZMNUOPT0

xIn = paramYBase_DataF & "ZMNUOPT0.txt"
intFile = FreeFile(0)
Open xIn For Input As #intFile
Do Until EOF(1)
    Line Input #intFile, xIn
    If Trim(xIn) <> "" Then
   
            xZMNUOPT0.MNUOPTCOD = CLng(Val(mId$(xIn, 1, 8)))
            xZMNUOPT0.MNUOPTCLI = mId$(xIn, 9, 7)
            xZMNUOPT0.MNUOPTLIB = mId$(xIn, 16, 35)
            xZMNUOPT0.MNUOPTENS = mId$(xIn, 51, 8)
            xZMNUOPT0.MNUOPTENT = mId$(xIn, 59, 8)
            xZMNUOPT0.MNUOPTSTR = mId$(xIn, 67, 1)
            xZMNUOPT0.MNUOPTARE = mId$(xIn, 68, 1)
            xZMNUOPT0.MNUOPTBAT = mId$(xIn, 69, 1)
            xZMNUOPT0.MNUOPTVAL = mId$(xIn, 70, 1)
            xZMNUOPT0.MNUOPTSUP = mId$(xIn, 71, 1)
            xZMNUOPT0.MNUOPTOIA = mId$(xIn, 72, 1)
            xZMNUOPT0.MNUOPTGES = mId$(xIn, 73, 1)

        V = adoZMNUOPT0_AddNew(rsSAB073Y, xZMNUOPT0)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_ZMNUOPT0"

    End If
    
Loop
Close intFile

End Sub

Public Sub cmdSelect_HISMVTP0_txt(lFile As String)
Dim V, xIn As String, intFile As Integer
Dim xYBIAMVT0 As typeYBIAMVT0
Dim K As Integer
Dim X As String, wJJ As Long, wMM As Long, wAAAA As Long
Dim X50 As String * 50
Dim Nb As Long
Dim wDevise As Integer, xCompte As String

Nb = 0
xIn = paramYBase_DataF & lFile
intFile = FreeFile(0)
Open xIn For Input As #intFile
Do Until EOF(1)
    Line Input #intFile, xIn
    If Trim(xIn) <> "" Then
            Nb = Nb + 1
            rsYBIAMVT0_Init xYBIAMVT0
            xYBIAMVT0.MOUVEMETA = 1
            xYBIAMVT0.MOUVEMPLA = 1
            xYBIAMVT0.MOUVEMSER = "00"
            xYBIAMVT0.MOUVEMSSE = "00"
            
            xYBIAMVT0.BIAMVTID = Nb

            K = 0
            X = CSV_Scan(xIn, K)
            X = CSV_Scan(xIn, K)
            wDevise = CInt(CSV_Scan(xIn, K))
            xCompte = CSV_Scan(xIn, K)
            If xCompte = "25272001014" Then
                Select Case wDevise
                    Case 8: xYBIAMVT0.MOUVEMCOM = "11178008001": xYBIAMVT0.COMPTEDEV = "DKK"
                    Case 400: xYBIAMVT0.MOUVEMCOM = "11178400001": xYBIAMVT0.COMPTEDEV = "USD"
                    Case 732: xYBIAMVT0.MOUVEMCOM = "11178732001": xYBIAMVT0.COMPTEDEV = "JPY"
                    Case 978: xYBIAMVT0.MOUVEMCOM = "11178978001": xYBIAMVT0.COMPTEDEV = "EUR"
                    Case Else: MsgBox xIn, vbInformation, "cmdSelect_HISMVTP0_txt"
                End Select
            End If
            X = CSV_Scan(xIn, K) '???
            
            xYBIAMVT0.MOUVEMOPE = CSV_Scan(xIn, K)
            X = CSV_Scan(xIn, K)
            xYBIAMVT0.MOUVEMMON = -(CCur(Val(CSV_Scan(xIn, K))))
            
            X = CSV_Scan(xIn, K)
            X = CSV_Scan(xIn, K)
            X = CSV_Scan(xIn, K)
            
            wJJ = CLng(CSV_Scan(xIn, K))
            wMM = CLng(CSV_Scan(xIn, K))
            wAAAA = CLng(CSV_Scan(xIn, K))
            xYBIAMVT0.MOUVEMDOP = wAAAA * 10000 + wMM * 100 + wJJ - 19000000
            
            
            wJJ = CLng(CSV_Scan(xIn, K))
            wMM = CLng(CSV_Scan(xIn, K))
            wAAAA = CLng(CSV_Scan(xIn, K))
            xYBIAMVT0.MOUVEMDVA = wAAAA * 10000 + wMM * 100 + wJJ - 19000000
            
            xYBIAMVT0.MOUVEMPIE = CLng(CSV_Scan(xIn, K))
            xYBIAMVT0.MOUVEMECR = CLng(CSV_Scan(xIn, K))
            X = CSV_Scan(xIn, K)
            
            X50 = ""
            X50 = CSV_Scan(xIn, K)
            xYBIAMVT0.LIBELLIB1 = mId$(X50, 1, 30)
            xYBIAMVT0.LIBELLIB2 = mId$(X50, 31, 20)
            
            X = CSV_Scan(xIn, K) 'bdfk
            X = CSV_Scan(xIn, K) 'bdfcode
            X = CSV_Scan(xIn, K) 'exo
            X = CSV_Scan(xIn, K) 'auto
            X = CSV_Scan(xIn, K) 'edi avis
            X = CSV_Scan(xIn, K) 'cpt comp
            X = CSV_Scan(xIn, K) '
            X = CSV_Scan(xIn, K)
            X = CSV_Scan(xIn, K)
            xYBIAMVT0.BIAMVTSD0 = -(CCur(Val(CSV_Scan(xIn, K)))) 'solde Veille
            
            wJJ = CLng(CSV_Scan(xIn, K))
            wMM = CLng(CSV_Scan(xIn, K))
            wAAAA = CLng(CSV_Scan(xIn, K))
            xYBIAMVT0.MOUVEMDTR = wAAAA * 10000 + wMM * 100 + wJJ - 19000000
            xYBIAMVT0.MOUVEMDCO = xYBIAMVT0.MOUVEMDTR
            

        V = adoYBIAMVT0_AddNew(rsSAB073Y, xYBIAMVT0)
        If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_YBIAMVT0"

    End If
    
Loop
Close intFile

End Sub

Public Sub cmdSelect_YCHQMON0()
Dim xSql As String
Dim xYCHQMON0 As typeYCHQMON0

xSql = "Delete * from YCHQMON0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from YCHQMON0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SABSPE & ".YCHQMON0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsYCHQMON0_GetBuffer(rsADO, xYCHQMON0)
    V = adoYCHQMON0_AddNew(rsSAB073Y, xYCHQMON0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_YCHQMON0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("YCHQMON0")
Set rsSAB073Y = Nothing
End Sub
Public Sub cmdSelect_YEUPMON0()
Dim xSql As String
Dim xYEUPMON0 As typeYEUPMON0

xSql = "Delete * from YEUPMON0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from YEUPMON0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SABSPE & ".YEUPMON0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsYEUPMON0_GetBuffer(rsADO, xYEUPMON0)
    V = adoYEUPMON0_AddNew(rsSAB073Y, xYEUPMON0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_YEUPMON0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("YEUPMON0")
Set rsSAB073Y = Nothing
End Sub

Public Sub cmdSelect_SideEUPLAB0()
'Uniquement appelé pour SAB073Y
Dim xSql As String
Dim xSideEUPLAB0 As typeSideEUPLAB0
Dim cnSQL_Server_BIA As New ADODB.Connection, rsSQL_Server_BIA As New ADODB.Recordset

xSql = "Delete * from SideEUPLAB0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

cnSQL_Server_BIA.Open paramODBC_DSN_SQL_Server_BIA

rsSAB073Y.Open "select * from SideEUPLAB0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from  sepa " & " order by EUPLABID"
Set rsSQL_Server_BIA = cnSQL_Server_BIA.Execute(xSql)

Do While Not rsSQL_Server_BIA.EOF

    Call rsSideEUPLAB0_GetBuffer(rsSQL_Server_BIA, xSideEUPLAB0)
    V = adoSideEUPLAB0_AddNew(rsSAB073Y, xSideEUPLAB0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_SideEUPLAB0"
    rsSQL_Server_BIA.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("SideEUPLAB0")
Set rsSAB073Y = Nothing
cnSQL_Server_BIA.Close
Set cnSQL_Server_BIA = Nothing

End Sub

Public Sub cmdSelect_YGUIMAD0()
Dim xSql As String
Dim xYGUIMAD0 As typeYGUIMAD0

xSql = "Delete * from YGUIMAD0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from YGUIMAD0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SABSPE & ".YGUIMAD0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsYGUIMAD0_GetBuffer(rsADO, xYGUIMAD0)
    V = adoYGUIMAD0_AddNew(rsSAB073Y, xYGUIMAD0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_YGUIMAD0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("YGUIMAD0")
Set rsSAB073Y = Nothing
End Sub

Public Sub cmdSelect_YROPINF0()
Dim xSql As String
Dim xYROPINF0 As typeYROPINF0

xSql = "Delete * from YROPINF0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from YROPINF0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SABSPE & ".YROPINF0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsYROPINF0_GetBuffer(rsADO, xYROPINF0)
    V = adoYROPINF0_AddNew(rsSAB073Y, xYROPINF0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_YROPINF0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("YROPINF0")
Set rsSAB073Y = Nothing
End Sub

Public Sub cmdSelect_YROPDOS0()
Dim xSql As String
Dim xYROPDOS0 As typeYROPDOS0

xSql = "Delete * from YROPDOS0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from YROPDOS0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SABSPE & ".YROPDOS0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsYROPDOS0_GetBuffer(rsADO, xYROPDOS0)
    V = adoYROPDOS0_AddNew(rsSAB073Y, xYROPDOS0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_YROPDOS0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("YROPDOS0")
Set rsSAB073Y = Nothing
End Sub

Public Sub cmdSelect_YBIASTO0()
Dim xSql As String
Dim xYBIASTO0 As typeYBIASTO0

xSql = "Delete * from YBIASTO0"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from YBIASTO0", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIASTO0"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsYBIASTO0_GetBuffer(rsADO, xYBIASTO0)
    V = adoYBIASTO0_AddNew(rsSAB073Y, xYBIASTO0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_YBIASTO0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("YBIASTO0")
Set rsSAB073Y = Nothing
End Sub
Public Sub cmdSelect_YBIAMON7()
Dim xSql As String
Dim xYBIAMON0 As typeYBIAMON0

xSql = "Delete * from YBIAMON7"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from YBIAMON7", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIAMON7"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsYBIAMON0_GetBuffer(rsADO, xYBIAMON0)
    V = adoYBIAMON0_AddNew(rsSAB073Y, xYBIAMON0)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_YBIAMON0"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("YBIAMON7")
Set rsSAB073Y = Nothing
End Sub

Public Sub cmdSelect_YBIARELV()
Dim xSql As String
Dim xYBIARELV As typeYBIARELV

xSql = "Delete * from YBIARELV"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from YBIARELV", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIARELV"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    Call rsYBIARELV_GetBuffer(rsADO, xYBIARELV)
    V = adoYBIARELV_AddNew(rsSAB073Y, xYBIARELV)
    If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " : " & "cmdSelect_YBIARELV"
    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("YBIARELV")
Set rsSAB073Y = Nothing
End Sub


Private Sub txtSelect_File_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub



Public Sub cmdSelect_ZSWIT001()
Dim X As String
X = "Etablissement integer , Agence integer , Service TEXT(2) , Sous_service TEXT(2) , Utilisateur TEXT(10)" _
& " , Reel TEXT(1) , Reception TEXT(1) , Saisie TEXT(1) , Modification TEXT(1) , Suppression TEXT(1) , Envoi TEXT(1) , Copie  TEXT(1)" _
& " , Validation  TEXT(1) , Montant1 currency , Montant2 currency , Type_Message TEXT(120)" _
& " , Ed_Detail  TEXT(1) , Ed_Page TEXT(1) , Ed_Gen TEXT(1) , Tri_Ref TEXT(1) , Tri_Valeur TEXT(1) , Tri_MT TEXT(1) , Tri_Dev TEXT(1)"
Call cmdSelect_Create("ZSWIT001", X)

Dim xSql As String

xSql = "Delete * from ZSWIT001"
Call FEU_ROUGE
Set rsSAB073Y = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "SAB073Y_temporisation de 5' (Delete *)"): DoEvents
Wait_SS 5

rsSAB073Y.Open "select * from ZSWIT001", paramODBC_DSN_SAB073Y, adOpenDynamic, adLockOptimistic

xSql = "select * from " & paramIBM_Library_SAB & ".ZSWITAB0 where SWITABNUM = 1 order by SWITABETA , SWITABARG"
Set rsADO = cnAdo.Execute(xSql)
Do While Not rsADO.EOF
    
    rsSAB073Y.AddNew
    rsSAB073Y("Etablissement") = rsADO("SWITABETA")
    X = rsADO("SWITABARG")
    rsSAB073Y("Agence") = Asc(mId$(X, 2, 1))
    rsSAB073Y("Service") = mId$(X, 3, 2)
    rsSAB073Y("Sous_service") = mId$(X, 5, 2)
    rsSAB073Y("Utilisateur") = mId$(X, 7, 10)
    
    X = rsADO("SWITABDON")
    rsSAB073Y("Reel") = mId$(X, 1, 1)
    rsSAB073Y("Reception") = mId$(X, 2, 1)
    rsSAB073Y("Ed_Detail") = mId$(X, 3, 1)
    rsSAB073Y("Ed_Page") = mId$(X, 4, 1)
    rsSAB073Y("Tri_Ref") = mId$(X, 5, 1)
    rsSAB073Y("Tri_Valeur") = mId$(X, 6, 1)
    rsSAB073Y("Tri_MT") = mId$(X, 7, 1)
    rsSAB073Y("Tri_Dev") = mId$(X, 8, 1)
    rsSAB073Y("Saisie") = mId$(X, 9, 1)
    rsSAB073Y("Modification") = mId$(X, 10, 1)
    rsSAB073Y("Suppression") = mId$(X, 11, 1)
    rsSAB073Y("Montant1") = CCur(convX2P(mId$(X, 12, 5)))
    rsSAB073Y("Validation") = mId$(X, 18, 1)
    rsSAB073Y("Type_Message") = mId$(X, 19, 120)
    rsSAB073Y("Envoi") = mId$(X, 139, 1)
    rsSAB073Y("Montant2") = CCur(convX2P(mId$(X, 140, 5)))
    rsSAB073Y("Ed_Gen") = mId$(X, 146, 1)
    rsSAB073Y("Copie") = mId$(X, 147, 1)
    
    rsSAB073Y.Update

    rsADO.MoveNext
Loop

rsSAB073Y.Close

Set rsADO = Nothing
Call traite_NEWDATE("ZSWIT001")
Set rsSAB073Y = Nothing


End Sub
Public Sub cmdSelect_Create(lTable_Name As String, lTable_Fields As String)
On Error Resume Next
Dim xSql As String
xSql = "DROP TABLE " & lTable_Name
cnSAB073Y.Execute xSql
On Error GoTo Error_Handler
xSql = "CREATE TABLE " & lTable_Name & " (" & lTable_Fields & ")"
cnSAB073Y.Execute xSql
GoTo Exit_sub

Error_Handler:
MsgBox Error, vbCritical, Me.Caption & " :  cmdSelect_Create"
Exit_sub:

End Sub

