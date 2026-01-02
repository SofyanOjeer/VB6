VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTC_limites 
   AutoRedraw      =   -1  'True
   Caption         =   "TC : Limites trésorerie PRE / EMP"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13560
   Icon            =   "TC_Limites.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9150
   ScaleWidth      =   13560
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
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
      TabPicture(0)   =   "TC_Limites.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Détail"
      TabPicture(1)   =   "TC_Limites.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraYBIASTO0"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "'"
      TabPicture(2)   =   "TC_Limites.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame fraYBIASTO0 
         Height          =   8025
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   13290
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
            TabIndex        =   8
            Top             =   120
            Width           =   11355
            Begin VB.TextBox txtSelect_TREOPECLI 
               Height          =   285
               Left            =   1800
               TabIndex        =   11
               Top             =   360
               Width           =   1845
            End
            Begin MSComCtl2.DTPicker txtSelect_AmjMin 
               Height          =   300
               Left            =   5760
               TabIndex        =   9
               Top             =   360
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
               Format          =   60555267
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtSelect_AmjMax 
               Height          =   300
               Left            =   7680
               TabIndex        =   10
               Top             =   360
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
               Format          =   60555267
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblSelect_TREOPECLI 
               Caption         =   "Client"
               Height          =   255
               Left            =   240
               TabIndex        =   12
               Top             =   360
               Width           =   1245
            End
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   645
            Left            =   11880
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6825
            Left            =   120
            TabIndex        =   5
            Top             =   1200
            Width           =   4080
            _ExtentX        =   7197
            _ExtentY        =   12039
            _Version        =   393216
            Rows            =   1
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   300
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
            FormatString    =   "<Racine|<Abrégé  |<P/E|                ||"
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
         Begin MSFlexGridLib.MSFlexGrid fgDossier 
            Height          =   3765
            Left            =   5040
            TabIndex        =   13
            Top             =   4320
            Width           =   7920
            _ExtentX        =   13970
            _ExtentY        =   6641
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
            FormatString    =   $"TC_Limites.frx":035E
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
      Picture         =   "TC_Limites.frx":03E9
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
Attribute VB_Name = "frmTC_limites"
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
Dim JRN_YTREOPE0_Aut As typeAuthorization
Dim curX1 As Currency, curX2 As Currency

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim cnADO As New ADODB.Connection
Dim rsADO As New ADODB.Recordset

Dim fgDossier_FormatString As String, fgDossier_K As Integer
Dim fgDossier_RowDisplay As Integer, fgDossier_RowClick As Integer, fgDossier_ColClick As Integer
Dim fgDossier_ColorClick As Long, fgDossier_ColorDisplay As Long
Dim fgDossier_Sort1 As Integer, fgDossier_Sort2 As Integer
Dim fgDossier_SortAD As Integer, fgDossier_Sort1_Old As Integer
Dim fgDossier_arrIndex As Integer
Dim blnfgDossier_DisplayLine As Boolean
Dim meYTREOPE0 As typeYTREOPE0, xYTREOPE0 As typeYTREOPE0
Dim arrYTREOPE0() As typeYTREOPE0
Dim selYTREOPE0() As typeYTREOPE0, selYTREOPE0_Nb As Long, selYTREOPE0_Max As Long


Dim meYCLIENA0 As typeYCLIENA0, xYCLIENA0 As typeYCLIENA0
Dim arrYCLIENA0() As typeYCLIENA0, arrClient_Nb As Long, arrClient_Max As Long
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

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
cmdPrint.Enabled = False

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgSelect_Display"
Set rsADO = Nothing

For I = 1 To arrClient_Nb
    xYTREOPE0 = arrYTREOPE0(I)
    X = "CLIENACLI =  '" & xYTREOPE0.TREOPECLI & "'"
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 where " & X
    
    blnOk = False
    blnDisplay = True
        Set rsADO = cnADO.Execute(xSQL)
        If Not rsADO.EOF Then
            V = srvYCLIENA0_GetBuffer_ODBC(rsADO, arrYCLIENA0(I))
            If Not IsNull(V) Then
                MsgBox V, vbCritical, "frmSAB_Stock.fgSelect_Display"
               '' Exit Sub
            Else
                fgSelect_DisplayLine (I)
            End If
        End If
    End If
    
Next I

fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
fgSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Comptes : " & arrClient_Nb): DoEvents
If fgSelect.Rows > 1 Then
    cmdPrint.Enabled = True
Else
    If chkSelect_Ecart.Value = "1" Then MsgBox "NEANT", vbInformation, "Tc_Limites"
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub


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
        For I = 0 To fgDossier_arrIndex
          fgDossier.Col = I: fgDossier.CellBackColor = lColor
        Next I
        fgDossier.Col = 0
    End If
End If

End Sub

Private Sub fgDossier_Display()
Dim xSQL As String
Dim intReturn As Integer
Dim curTotal As Currency, nbTotal As Long, curSolde As Currency
fgDossier_Reset
fgDossier.Rows = 1
fgDossier.FormatString = fgDossier_FormatString

ReDim selYTREOPE0(101)
selYTREOPE0_Max = 99: selYTREOPE0_Nb = 0

Set rsADO = Nothing
libYTREOPE0_Diff = ""
curTotal = 0: nbTotal = 0
If blnJPL Then Exit Sub
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YTREOPE0 where " _
     & "YSTOPCI = '" & xYTREOPE0.YSTOPCI & "'" _
     & "AND YSTODEV = '" & xYTREOPE0.YSTODEV & "'" _
     & "AND YSTOCLI = " & xYTREOPE0.YSTOCLI

Set rsADO = cnADO.Execute(xSQL)

Do While Not rsADO.EOF
    V = srvYTREOPE0_GetBuffer_ODBC(rsADO, xYTREOPE0)
    If Not IsNull(V) Then
        MsgBox V, vbCritical, "frmSAB_Balance.SQL_ODBC"
        Exit Sub
    Else
        nbTotal = nbTotal + 1
        fgDossier_DisplayLine nbTotal
        curTotal = curTotal + xYTREOPE0.YSTOMON
        
        selYTREOPE0_Nb = selYTREOPE0_Nb + 1
        If selYTREOPE0_Nb > selYTREOPE0_Max Then
            selYTREOPE0_Max = selYTREOPE0_Max + 100
            ReDim Preserve selYTREOPE0(selYTREOPE0_Max)
        End If
        
       selYTREOPE0(selYTREOPE0_Nb) = xYTREOPE0

    End If
    rsADO.MoveNext
Loop
libYTREOPE0_Total = nbTotal & " dossiers : " & xYCLIENA0.COMPTEDEV
libYTREOPE0_Solde = "Solde " & xYCLIENA0.COMPTEDEV & " au " & dateIBM10(YBIATAB0_DATE_CPT_J, True)

libYTREOPE0_YSTOMON = Format$(curTotal, "### ### ### ###.00")
curSolde = Abs(xYCLIENA0.SOLDECEN)
libYTREOPE0_SOLDECEN = Format$(curSolde, "### ### ### ###.00")

If curTotal = curSolde Then
    libYTREOPE0_Diff = ""
Else
    libYTREOPE0_Diff.ForeColor = vbRed
    libYTREOPE0_Diff = Format$(curTotal - curSolde, "### ### ### ###.00")
End If


'libYTREOPE0 = nbTotal & " dossiers, " & Format$(curTotal, "### ### ### ###.00") & " " & xYCLIENA0.COMPTEDEV
fgDossier.Visible = True
fgDossier_Sort1 = -1
End Sub

Public Sub fgDossier_DisplayLine(lIndex As Long)
On Error Resume Next
If chkSelect_YSTOMON = "0" And xYTREOPE0.YSTOMON = 0 Then Exit Sub
fgDossier.Rows = fgDossier.Rows + 1
fgDossier.Row = fgDossier.Rows - 1
fgDossier.Col = 0: fgDossier.Text = xYTREOPE0.YSTOAPP & " " & xYTREOPE0.YSTOOPE & " " & xYTREOPE0.YSTONAT & " " & xYTREOPE0.YSTONUM
fgDossier.Col = 1: fgDossier.Text = Format$(xYTREOPE0.YSTOMON, "### ### ### ###.00")
fgDossier.Col = 2: fgDossier.Text = xYTREOPE0.YSTODEV
fgDossier.Col = 3: fgDossier.Text = dateImp10(xYTREOPE0.YSTODEB)
fgDossier.Col = 4: fgDossier.Text = dateImp10(xYTREOPE0.YSTOFIN)
fgDossier.Col = 5: fgDossier.Text = xYTREOPE0.YSTOCLI
fgDossier.Col = 6: fgDossier.Text = Trim(xYTREOPE0.YSTOPCI)
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

Public Sub fgSelect_DisplayLine(lIndex As Integer)
On Error Resume Next

Dim curX As Currency
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1
fgSelect.Col = 0: fgSelect.Text = xYCLIENA0.CLIENACLI
fgSelect.Col = 1: fgSelect.Text = xYCLIENA0.CLIENASIG
fgSelect.Col = 2: fgSelect.Text = xYTREOPE0.TREOPEOPR

fgSelect.Col = 3: fgSelect.Text = Format$(xYTREOPE0.TREOPEMNT, "### ### ### ###.00")

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


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), JRN_YTREOPE0_Aut)
Form_Init


End Sub


Public Sub Form_Init()
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistent", vbCritical, "frmYTREOPE0.param_init"
    Unload Me
Else
    lstErr.Clear
End If

blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fgDossier_FormatString = fgDossier.FormatString
fgSelect.Enabled = True
tableYBase_Open
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
recYTREOPE0_Init meYTREOPE0
xYTREOPE0 = meYTREOPE0
fraSelect_Options.Enabled = False
cmdSelect_Ok_Click
blnControl = True



End Sub


Public Function param_Init()

param_Init = Null
Call lstErr_Clear(lstErr, cmdContext, "param_Init"): DoEvents

fgSelect.Visible = False

fgSelect.Visible = True


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

Private Sub txtselect_cdodosdos_GotFocus()
txt_GotFocus txtSelect_CDODOSDOS

End Sub

Private Sub txtselect_cdodosdos_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtselect_cdodosdos_LostFocus()
txt_LostFocus txtSelect_CDODOSDOS

End Sub

Private Sub txtselect_amjmin_GotFocus()
txt_GotFocus txtSelect_AmjMin

End Sub

Private Sub txtselect_amjmin_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtselect_amjmin_LostFocus()
txt_LostFocus txtSelect_AmjMin

End Sub

Private Sub txtselect_amjmax_GotFocus()
txt_GotFocus txtSelect_AmjMax

End Sub

Private Sub txtselect_amjmax_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtselect_amjmax_LostFocus()
txt_LostFocus txtSelect_AmjMax

End Sub










Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
Dim blnOk As Boolean
Dim I As Integer, nbDossier As Long

ReDim arrYTREOPE0(101)
arrClient_Max = 50: arrClient_Nb = 0
blnOk = False
nbDossier = 0

xWhere = " where TREOPESTA <= '3'"

X = Trim(txtSelect_TREOPECLI)

If X <> "" Then
    If xWhere = "" Then
        xAnd = " where "
    Else
        xAnd = " and "
    End If
    xWhere = xWhere & xAnd & "TREOPECLI = '00" & X & "'"
End If

xSQL = "select * from " & paramIBM_Library_SAB & ".ZTREOPE0" & xWhere & " order by TREOPECLI, TREOPEOPR"
If blnJPL Then Exit Sub
Set rsADO = cnADO.Execute(xSQL)

Do While Not rsADO.EOF
    nbDossier = nbDossier + 1
    V = srvYTREOPE0_GetBuffer_ODBC(rsADO, xYTREOPE0)
   
   If Not IsNull(V) Then
        MsgBox V, vbCritical, "frmSAB_Stock.cmdSelect_Ok_Click"
        Exit Sub
    Else
        If Not blnOk Then
            blnOk = True
            arrClient_Nb = 1
            arrYCLIENA0(1).CLIENACLI = xYTREOPE0.TREOPECLI
        Else
            If xYTREOPE0.TREOPECLI = arrYTREOPE0(arrClient_Nb).TREOPECLI _
            And xYTREOPE0.TREOPEOPR = arrYTREOPE0(arrClient_Nb).TREOPEOPR Then
                arrYTREOPE0(arrClient_Nb).TREOPEMNT = arrYTREOPE0(arrClient_Nb).TREOPEMNT + xYTREOPE0.TREOPEMNT
        
            Else
                arrClient_Nb = arrClient_Nb + 1
                If arrClient_Nb > arrClient_Max Then
                    arrClient_Max = arrClient_Max + 50
                    ReDim Preserve arrYTREOPE0(arrClient_Max)
                End If
                
                arrYTREOPE0(arrClient_Nb) = xYTREOPE0
            End If
        End If
    End If
    rsADO.MoveNext
Loop
Call lstErr_AddItem(lstErr, cmdContext, "Lignes d'encours : " & nbDossier): DoEvents

ReDim arrYCLIENA0(arrClient_Max)
fgSelect_Display
End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAb_Stock_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
fgDossier.Clear
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


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim K As Long
On Error Resume Next
If Y <= fgSelect.RowHeightMin Then
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
        xYTREOPE0 = arrYTREOPE0(K)
        xYCLIENA0 = arrYCLIENA0(K)
        fgDossier_Display
        SSTab1.Tab = 1
        
   End If
End If
End Sub

Private Sub fgDossier_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim K As Long
On Error Resume Next
If Y <= fgDossier.RowHeightMin Then
    Select Case fgDossier.Col
        Case 0: fgDossier_Sort1 = 0: fgDossier_Sort2 = 1: fgDossier_Sort
        Case 1:  fgDossier_Sort1 = 1: fgDossier_Sort2 = 1: fgDossier_Sort
        Case 2: fgDossier_Sort1 = 2: fgDossier_Sort2 = 2: fgDossier_Sort
        Case 3: fgDossier_Sort1 = 3: fgDossier_Sort2 = 3: fgDossier_Sort
        Case 4: fgDossier_Sort1 = 4: fgDossier_Sort2 = 4: fgDossier_Sort
    End Select
Else
    If fgDossier.Rows > 1 Then
        Call fgDossier_Color(fgDossier_RowClick, MouseMoveUsr.BackColor, fgDossier_ColorClick)
        fgDossier.Col = fgDossier_arrIndex:  K = CLng(fgDossier.Text)
        xYTREOPE0 = selYTREOPE0(K)
        srvYTREOPE0_ElpDisplay xYTREOPE0

   End If
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'If blnControl Then
    cnADO.Close
    Set cnADO = Nothing
'End If
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

cnADO.Open paramODBC_DSN_SAB 'JRN
Exit Sub

Error_Handler:
blnControl = False
MsgBox Error
End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    Case 1: fgDossier.LeftCol = 0
End Select
End Sub


Public Sub cmdPrint_Ok(lFct As String)
fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Liste : " & fgSelect.Rows - 1)

prtSAB_Stock_Monitor lFct, fgSelect, arrYTREOPE0(), arrYCLIENA0(), arrClient_Nb
fgSelect.Visible = True
Me.Show
End Sub
