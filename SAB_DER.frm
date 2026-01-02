VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSAB_DER 
   AutoRedraw      =   -1  'True
   Caption         =   "SAb_DER"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13560
   Icon            =   "SAB_DER.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9150
   ScaleWidth      =   13560
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6720
      TabIndex        =   3
      Top             =   0
      Width           =   6252
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
      TabCaption(0)   =   "SAB : Groupes 7*** 8*** 9***"
      TabPicture(0)   =   "SAB_DER.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "....."
      TabPicture(1)   =   "SAB_DER.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTab1"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraTab1 
         Height          =   7932
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   13212
      End
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13290
         Begin VB.CommandButton cmdUpdate 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Créer un groupe"
            Height          =   612
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   7320
            Visible         =   0   'False
            Width           =   1692
         End
         Begin VB.Frame fraUpdate 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Création d'un groupe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1092
            Left            =   2520
            TabIndex        =   11
            Top             =   7080
            Visible         =   0   'False
            Width           =   10452
            Begin VB.TextBox txtUpdate_CLIENARA1 
               Height          =   288
               Left            =   1800
               MaxLength       =   32
               TabIndex        =   18
               Text            =   "raison sociale"
               Top             =   720
               Width           =   5052
            End
            Begin VB.TextBox txtUpdate_CLIENASIG 
               Height          =   288
               Left            =   1800
               MaxLength       =   12
               TabIndex        =   17
               Text            =   "sigle "
               Top             =   360
               Width           =   3252
            End
            Begin VB.CommandButton cmdUpdate_Quit 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Abandonner"
               Height          =   525
               Left            =   8760
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   240
               Width           =   1575
            End
            Begin VB.CommandButton cmdUpdate_Ok 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Enregistrer"
               Height          =   525
               Left            =   7080
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label lblUpdate_CLIENARA1 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Raison sociale"
               Height          =   252
               Left            =   120
               TabIndex        =   16
               Top             =   720
               Width           =   1500
            End
            Begin VB.Label lblUpdate_CLIENASIG 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Sigle"
               Height          =   252
               Left            =   120
               TabIndex        =   15
               Top             =   360
               Width           =   1500
            End
         End
         Begin VB.ComboBox cboSelect_SQL 
            Height          =   315
            Left            =   9480
            TabIndex        =   9
            Text            =   "cboSelect_SQL"
            Top             =   120
            Width           =   3615
         End
         Begin VB.Frame fraSelect_Options 
            Height          =   1005
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   6075
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Exécuter la requête"
            Height          =   525
            Left            =   11040
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   600
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   5868
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   12840
            _ExtentX        =   22648
            _ExtentY        =   10345
            _Version        =   393216
            Rows            =   1
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   350
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
            FormatString    =   "<Id        |<Sigle                            |<Raison sociale                                                             ||"
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
      Picture         =   "SAB_DER.frx":047A
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
Attribute VB_Name = "frmSAB_DER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit

Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String, currentError As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim SAB_DER_Aut As typeAuthorization
Dim curX1 As Currency, curX2 As Currency
Dim blnAuto As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim blnTransaction As Boolean
Dim cmdSelect_SQL_K As String, cmdSelect_SQL_Where As String, cmdSelect_SQL_RA1 As String
'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long

Dim xZCLIENA0 As typeZCLIENA0, newZCLIENA0 As typeZCLIENA0, oldZCLIENA0 As typeZCLIENA0
Dim arrZCLIENA0() As typeZCLIENA0, arrZCLIENA0_Nb As Long, arrZCLIENA0_Max As Long, arrZCLIENA0_Index As Long

Dim xZCLIENB0 As typeZCLIENB0, newZCLIENB0 As typeZCLIENB0, oldZCLIENB0 As typeZCLIENB0
Dim xZCLINPR0 As typeZCLINPR0, newZCLINPR0 As typeZCLINPR0, oldZCLINPR0 As typeZCLINPR0
Dim xZADRESS0 As typeZADRESS0, newZADRESS0 As typeZADRESS0, oldZADRESS0 As typeZADRESS0

Dim wCLIENACLI As Long, wCLIENACLI_Max As Long, wCLIENACLI_Min As Long

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
Private Sub fgSelect_Display_1()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = "fgselect_Display"
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
    
For I = 1 To arrZCLIENA0_Nb
         
    xZCLIENA0 = arrZCLIENA0(I)
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine_1 I
Next I

fgSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Nb de comptes : " & arrZCLIENA0_Nb): DoEvents
'If fgSelect.Rows > 1 Then
'    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
'End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub arrZCLIENA0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
Dim wCli As Long
On Error GoTo Error_Handler
ReDim arrZCLIENA0(101)
arrZCLIENA0_Max = 100: arrZCLIENA0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 " & xWhere & " order by CLIENACLI desc"
Set rsSab = cnsab.Execute(xSql)
wCLIENACLI = 0

Do While Not rsSab.EOF
    V = rsZCLIENA0_GetBuffer(rsSab, xZCLIENA0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmSAB_DER.fgselect_Display"
        '' Exit Sub
     Else
         arrZCLIENA0_Nb = arrZCLIENA0_Nb + 1
         If arrZCLIENA0_Nb > arrZCLIENA0_Max Then
             arrZCLIENA0_Max = arrZCLIENA0_Max + 50
             ReDim Preserve arrZCLIENA0(arrZCLIENA0_Max)
         End If
         wCli = CLng(xZCLIENA0.CLIENACLI)
         If wCLIENACLI < wCli Then wCLIENACLI = wCli

         arrZCLIENA0(arrZCLIENA0_Nb) = xZCLIENA0
    End If
    rsSab.MoveNext

Loop

If wCLIENACLI = 0 Then
    wCLIENACLI = wCLIENACLI_Min
    V = "!!!!  wCLIENACLI = " & wCLIENACLI_Min
    cmdUpdate.Visible = False
    GoTo Error_MsgBox
End If
If wCLIENACLI >= wCLIENACLI_Max Then
    V = "erreur  wCLIENACLI > " & wCLIENACLI_Max
    cmdUpdate.Visible = False
    GoTo Error_MsgBox
End If

wCLIENACLI = wCLIENACLI + 1

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub fgSelect_DisplayLine_1(lIndex As Long)
On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = xZCLIENA0.CLIENACLI
fgSelect.Col = 1: fgSelect.Text = xZCLIENA0.CLIENASIG
fgSelect.Col = 2: fgSelect.Text = xZCLIENA0.CLIENARA1

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


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), SAB_DER_Aut)
Form_Init

Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case "@SAB_DER": blnAuto = True
                    
                    Unload Me
           
    Case Else: blnAuto = False
                    

End Select


End Sub


Public Sub Form_Init()
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
blnControl = False

fgSelect_FormatString = fgSelect.FormatString

fgSelect.Visible = False
cmdUpdate.Visible = False
fraUpdate.Visible = False

'_____________________________________________________________


cboSelect_SQL.Clear
cboSelect_SQL.AddItem "6 - groupes de sécurité financière"
cboSelect_SQL.AddItem "7 - groupes de trésorerie"
cboSelect_SQL.AddItem "8 - groupes grands risques"
cboSelect_SQL.AddItem "9 - groupes économiques"

cboSelect_SQL.ListIndex = 0


fgSelect.Enabled = True
cmdReset

Me.Enabled = True
Me.MousePointer = 0
End Sub


Private Sub cboSelect_SQL_Click()
cmdSelect_SQL_K = mId$(cboSelect_SQL, 1, 1)

fraSelect_Options.Enabled = True
cmdSelect_Ok_Click
End Sub



'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
'lstErr.Visible = False
currentAction = ""
fraSelect_Options.Enabled = True
'cmdSelect_Ok_Click



blnControl = True



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
            If fgSelect.Rows > 1 Then cmdPrint_Ok
                
    Case 1:
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_SQL()
Dim V, fsoFile As File
Dim X As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL"
Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents


Call arrZCLIENA0_SQL(cmdSelect_SQL_Where)

fgSelect_Display_1

cmdUpdate.Visible = SAB_DER_Aut.Saisir

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Function cmdUpdate_Transaction()
Dim V, X As String, xSql As String
Dim NB As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdUpdate_Transaction"
'-------------------------------------------------------
cmdUpdate_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
V = sqlZCLIENA0_Insert(newZCLIENA0)
If Not IsNull(V) Then GoTo Error_MsgBox
V = sqlZCLIENB0_Insert(newZCLIENB0)
If Not IsNull(V) Then GoTo Error_MsgBox
V = sqlZCLINPR0_Insert(newZCLINPR0)
If Not IsNull(V) Then GoTo Error_MsgBox
V = sqlZADRESS0_Insert(newZADRESS0)
If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_Sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
    cmdUpdate_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, NB As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAB_DER_cmdSelect_Ok ........"): DoEvents
cmdUpdate.Visible = False
fraUpdate.Visible = False

fgSelect.Clear
If blnOk Then
    cmdSelect_Ok.Caption = "Options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = False
    Select Case cmdSelect_SQL_K
        Case "6":   cmdSelect_SQL_RA1 = "GRP SFI "
                    wCLIENACLI_Min = 6000: wCLIENACLI_Max = 6999
                    cmdSelect_SQL_Where = "where CLIENACLI >= '0006000' and  CLIENACLI <= '0006999'"
                    cmdSelect_SQL
        Case "7":   cmdSelect_SQL_RA1 = "GRP TRE "
                    wCLIENACLI_Min = 7000: wCLIENACLI_Max = 7999
                    cmdSelect_SQL_Where = "where CLIENACLI >= '0007000' and  CLIENACLI <= '0007999'"
                    cmdSelect_SQL
        Case "8":   cmdSelect_SQL_RA1 = "GRP RSQ "
                    wCLIENACLI_Min = 8000: wCLIENACLI_Max = 8999
                    cmdSelect_SQL_Where = "where CLIENACLI >= '0008000' and  CLIENACLI <= '0008999'"
                    cmdSelect_SQL
       Case "9":    cmdSelect_SQL_RA1 = "GRP ECO "
                    wCLIENACLI_Min = 9000: wCLIENACLI_Max = 9999
                    cmdSelect_SQL_Where = "where CLIENACLI >= '0009000' and  CLIENACLI <= '0009999'"
                    cmdSelect_SQL
    End Select
Else
    cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_DER_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAB_DER_cmdSelect_Ok ........"): DoEvents
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_DER_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdUpdate_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
rsZCLIENA0_Init newZCLIENA0
newZCLIENA0.CLIENACLI = Format$(wCLIENACLI, "0000000")
fraUpdate.Caption = "Création du groupe : " & wCLIENACLI
txtUpdate_CLIENASIG = cmdSelect_SQL_RA1 & wCLIENACLI
txtUpdate_CLIENARA1 = cmdSelect_SQL_RA1 & wCLIENACLI
cmdUpdate.Visible = False
fraUpdate.Visible = True

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdUpdate_Ok_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents

If IsNull(fraUpdate_Control) Then
    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents
    V = cmdUpdate_Transaction
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        cmdUpdate.Visible = SAB_DER_Aut.Saisir
        fraUpdate.Visible = False

        cmdSelect_SQL
    Else
        MsgBox V, vbCritical, Me.Name & " : cmdUpdate_Ok"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Function fraUpdate_Control()
Dim blnUpdate_Control As Boolean
Dim X As String
blnUpdate_Control = True
Call lstErr_AddItem(lstErr, cmdContext, ">_________Contrôle des données "): DoEvents

X = Trim(txtUpdate_CLIENASIG)
If X = "" Then
    blnUpdate_Control = False
    txtUpdate_CLIENASIG.BackColor = errUsr.BackColor
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le sigle")
Else
    txtUpdate_CLIENASIG.BackColor = txtUsr.BackColor
End If
newZCLIENA0.CLIENASIG = X


X = Trim(txtUpdate_CLIENARA1)
If X = "" Then
    blnUpdate_Control = False
    txtUpdate_CLIENARA1.BackColor = errUsr.BackColor
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le sigle")
Else
    txtUpdate_CLIENARA1.BackColor = txtUsr.BackColor
End If
newZCLIENA0.CLIENARA1 = X

newZCLIENA0.CLIENAETB = currentZMNURUT0.MNURUTETB
newZCLIENA0.CLIENAAGE = 1
newZCLIENA0.CLIENAETA = "GRPM"
newZCLIENA0.CLIENASRN = "NEANT"
newZCLIENA0.CLIENANAT = "FR"
newZCLIENA0.CLIENARSD = "FR"
newZCLIENA0.CLIENARES = "R50"
newZCLIENA0.CLIENAECO = "C01"
newZCLIENA0.CLIENACAT = "AUT"
newZCLIENA0.CLIENACHQ = "N"
newZCLIENA0.CLIENAENT = "000"
newZCLIENA0.CLIENAMES = "1"
newZCLIENA0.CLIENADOU = "N"
newZCLIENA0.CLIENACOL = "2"
newZCLIENA0.CLIENACRE = valDSys - 19000000

rsZCLINPR0_Init newZCLINPR0
newZCLINPR0.CLINPRETA = newZCLIENA0.CLIENAETB
newZCLINPR0.CLINPRCLI = newZCLIENA0.CLIENACLI
newZCLINPR0.CLINPRTYP = "1"
newZCLINPR0.CLINPRNUM = newZCLIENA0.CLIENASRN

rsZCLIENB0_Init newZCLIENB0
newZCLIENB0.CLIENBETB = newZCLIENA0.CLIENAETB
newZCLIENB0.CLIENBCLI = newZCLIENA0.CLIENACLI
newZCLIENB0.CLIENBAF1 = "00"
newZCLIENB0.CLIENBAF2 = "00"
newZCLIENB0.CLIENBAF3 = "00"
newZCLIENB0.CLIENBCTL = "N"
newZCLIENB0.CLIENBTOP = "1"

rsZADRESS0_Init newZADRESS0
newZADRESS0.ADRESSETA = newZCLIENA0.CLIENAETB
newZADRESS0.ADRESSTYP = "1"
newZADRESS0.ADRESSNUM = " " & newZCLIENA0.CLIENACLI
newZADRESS0.ADRESSAD1 = "DEPARTEMENT DES RISQUES"
newZADRESS0.ADRESSCOP = "75008"
newZADRESS0.ADRESSVIL = "PARIS"

If blnUpdate_Control Then
    fraUpdate_Control = Null
Else
    fraUpdate_Control = "?_________fraUpdate_Control"
End If
End Function


Private Sub cmdUpdate_Quit_Click()
cmdUpdate.Visible = SAB_DER_Aut.Saisir
fraUpdate.Visible = False

End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim K As Long
On Error Resume Next
If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
        Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  arrZCLIENA0_Index = CLng(fgSelect.Text)
        fgSelect.LeftCol = 0

   End If
End If
fgSelect.LeftCol = 0
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
    fgSelect.Col = fgSelect_arrIndex:
Else
    SendKeys "{TAB}"
End If
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
Dim V
Dim xName  As String, xMemo As String
On Error GoTo Error_Handler

mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False


Exit Sub

Error_Handler:

blnControl = False
If Not blnAuto Then MsgBox Error
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
cmdPrint_Ok '"D "
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_Liste_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Ok '"L "
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then cmdSelect_Ok.SetFocus

End Sub














Private Sub SSTab1_GotFocus()
Select Case SSTab1.Tab
    Case 0: fgSelect.LeftCol = 0
End Select
End Sub


Public Sub cmdPrint_Ok()
Dim iRow As Integer, K As Integer, I As Integer
Dim blnOk As Boolean

fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Etat : " & fgSelect.Rows - 1)

fgSelect.Visible = True
Me.Show
End Sub


Public Sub Error_Route(V)

currentError = CStr(V) & "             ( " & Me.Name & " ~ " & App_Debug & " )"
If blnAuto Then
  '  Call cmdSendMail_Alerte(Me.Name & " ~ " & App_Debug, CStr(V))
Else
    MsgBox V, vbCritical, Me.Name & " ~ " & App_Debug
End If

End Sub




Private Sub txtUpdate_CLIENARA1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtUpdate_CLIENASIG_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


