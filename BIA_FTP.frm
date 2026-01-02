VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBIA_FTP 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_CRE : Crédits "
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13560
   Icon            =   "BIA_FTP.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9150
   ScaleWidth      =   13560
   Begin VB.Frame fraSelect_Options 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5970
      Left            =   4680
      TabIndex        =   11
      Top             =   2280
      Width           =   7575
      Begin MSComCtl2.DTPicker txtAmjMax 
         Height          =   300
         Left            =   6120
         TabIndex        =   12
         Top             =   1200
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
         Format          =   22806531
         CurrentDate     =   36299
         MaxDate         =   401768
         MinDate         =   -328351
      End
      Begin MSComCtl2.DTPicker txtAmjMin 
         Height          =   300
         Left            =   4440
         TabIndex        =   13
         Top             =   1200
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
         Format          =   22806531
         CurrentDate     =   36299
         MaxDate         =   401768
         MinDate         =   -328351
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   7320
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7320
         Y1              =   960
         Y2              =   960
      End
   End
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
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Rechercher"
      TabPicture(0)   =   "BIA_FTP.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Dossier"
      TabPicture(1)   =   "BIA_FTP.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDossier"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "X"
      TabPicture(2)   =   "BIA_FTP.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Détail fichiers"
      TabPicture(3)   =   "BIA_FTP.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fgDossier"
      Tab(3).ControlCount=   1
      Begin VB.Frame fraDossier 
         Height          =   8145
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   13260
         Begin VB.Frame fraDossier_Saisie 
            Height          =   615
            Left            =   7320
            TabIndex        =   8
            Top             =   7440
            Width           =   5895
         End
      End
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   135
         TabIndex        =   3
         Top             =   330
         Width           =   13290
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   525
            Left            =   11760
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox txtSelect 
            Height          =   285
            Left            =   135
            TabIndex        =   6
            Top             =   240
            Width           =   1230
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7185
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   13080
            _ExtentX        =   23072
            _ExtentY        =   12674
            _Version        =   393216
            Rows            =   1
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   200
            BackColor       =   14737632
            ForeColor       =   4210688
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            TextStyle       =   4
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   "Application     |Fichier         |Libellé                           |Destination     |     ||"
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
      Begin MSFlexGridLib.MSFlexGrid fgDossier 
         Height          =   7755
         Left            =   -74880
         TabIndex        =   9
         Top             =   600
         Width           =   13275
         _ExtentX        =   23416
         _ExtentY        =   13679
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         RowHeightMin    =   200
         BackColor       =   14737632
         ForeColor       =   4210688
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12648384
         BackColorBkg    =   14737632
         AllowBigSelection=   0   'False
         TextStyle       =   4
         TextStyleFixed  =   4
         FocusRect       =   2
         HighLight       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   3
         FormatString    =   $"BIA_FTP.frx":037A
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
      Picture         =   "BIA_FTP.frx":0434
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContext_Auto 
         Caption         =   "Auto :impression MAD, Avis d'échéance"
      End
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
      Begin VB.Menu mnuPrint0_Avis 
         Caption         =   "Imprimer les avis"
      End
   End
End
Attribute VB_Name = "frmBIA_FTP"
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
Dim BIA_FTP_Aut As typeAuthorization
Dim xAmjMin As String, xAmjMax As String
Dim IbmAmjMin As String, IbmAmjMax As String

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim xElpTable As typeElpTable

Dim cnADO As New ADODB.Connection
Dim rsADO As New ADODB.Recordset
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
Dim Nb As Long
Dim xSQL As String
Dim xW As String
On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgSelect_Display"

    Set rsADO = Nothing
    
    xSQL = "select CREPREDOS , CREPREPRE ,CREPRENAT , CREPREDEV , CREPREMON  , CREPRECTA  , CREPREOUV from ZCREPRE0"
   Set rsADO = cnADO.Execute(xSQL)  '$2003.11.04

    Do While Not rsADO.EOF
        'Call srvYCREPRE0_GetBuffer_ODBC(rsADO, meYCREPRE0)
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            fgSelect.Col = 0: fgSelect.Text = rsADO("CREPREDOS")
           fgSelect.CellForeColor = vbBlue
            fgSelect.Col = 1: fgSelect.Text = rsADO("CREPREPRE")
            fgSelect.Col = 2: fgSelect.Text = rsADO("CREPRENAT")
   
            fgSelect.Col = 3: fgSelect.Text = Format$(rsADO("CREPREMON"), "### ### ### ##0.00") & rsADO("CREPREDEV")
            fgSelect.CellForeColor = vbBlue
            
            fgSelect.Col = 4: fgSelect.Text = rsADO("CREPRECTA")
            fgSelect.Col = 5: fgSelect.Text = dateIBM10(rsADO("CREPREOUV"), True)
  
        rsADO.MoveNext
    Loop
    '$2003.11.04  rsADO.Close
fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
fgSelect.Visible = True
fgSelect.TopRow = fgSelect.Rows - 12
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub fgSelect_DisplayLine()
On Error Resume Next
End Sub


Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 6
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

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), BIA_FTP_Aut)

'blnSetfocus = True
Form_Init


End Sub


Public Sub Form_Init()
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistent", vbCritical, "frmYCREPRE0.param_init"
    Unload Me
Else
    lstErr.Clear
End If

blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgDossier_FormatString = fgDossier.FormatString
fgDossier.Enabled = True
Call DTPicker_Set(txtAmjMax, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtAmjMin, YBIATAB0_DATE_CPT_J)

tableYBase_Open
cmdReset
Me.Enabled = True
Me.MousePointer = 0
End Sub
Public Sub txtAMJ_Control()
Call DTPicker_Control(txtAmjMax, xAmjMax)
IbmAmjMax = dateIBM(xAmjMax)
Call DTPicker_Control(txtAmjMin, xAmjMin)
IbmAmjMin = dateIBM(xAmjMin)
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

SSTab1.Tab = 0

recYCREPRE0_Init meYCREPRE0
xYCREPRE0 = meYCREPRE0
ReDim meYBIACRE.YCREPRE0(1): meYBIACRE.YCREPRE0(1) = meYCREPRE0


fraDossier.Enabled = False
lstDossier_Contact.ForeColor = vbMagenta

fraDossier_Info.Enabled = False   ' La frame n'est que affichage d'informations
cmdPrint.Enabled = BIA_FTP_Aut.Saisir

fraSelect_Options.Visible = False
fraSelect_Options.Top = 1560
fraSelect_Options.Left = 5600

cmdSelect_Ok_Click
blnControl = True

End Sub


Public Function param_Init()
Dim K As Integer, K1 As Integer, X As String
Dim wText As String
Dim V

param_Init = Null


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

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Dim prtI As Integer
Dim I As Long, iLen As Long


If fgSelect.Rows > 1 Then
        Select Case SSTab1.Tab
            Case 0:
                If optSelect_Avis Then
                    Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
                End If
            Case 1: cmdPrint_lstDossier_YCREAVI0
                    cmdPrint_lstDossier_YCREBIS0
            Case Else:
                    MsgBox "Impression non gérée pour cet onglet", vbCritical, "frmBIA_FTP.cmdPrint"
                
        End Select
End If

End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean

blnOk = fraSelect_Options.Visible
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAb_Stock_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
lstDossier_YCREAVI0.Clear
lstDossier_YCREBIS0.Clear
If blnOk Then
    cmdSelect_Ok.Caption = "Options"
    cmdSelect_Ok.BackColor = &HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Visible = False
    cmsSelect_SQL
Else
    cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Visible = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAb_Stock_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub fgDossier_Click()
fgDossier.LeftCol = 0

End Sub

Private Sub fgDossier_LeaveCell()
On Error Resume Next
fgDossier.CellBackColor = &HE0E0E0
End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wK1 As Long, wK2 As Long
On Error Resume Next
If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_SortX fgSelect_Sort1
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
       ' fgSelect.Col = fgSelect_arrIndex: wK1 = fgSelect.Text
        fgSelect.Col = 0: wK1 = CLng(Val(fgSelect.Text))
        fgSelect.Col = 1: wK2 = CLng(Val(fgSelect.Text))
        cmdSelect wK1, wK2
   End If
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
If blnControl Then
    cnADO.Close
    Set cnADO = Nothing
End If
End Sub

Private Sub mnuContext_Auto_Click()
chkSelect_Print = "1"
optSelect_Avis_Echéance = True
fraSelect_Options.Visible = True
cmdSelect_Ok_Click
fraSelect_Options.Visible = True
optSelect_MAD = True
cmdSelect_Ok_Click

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
    If optSelect_Dossier Then
        fgSelect.Row = fgSelect.TopRow
        fgSelect.Col = fgSelect_arrIndex: ' wK1 = fgSelect.Text
        'cmdSelect txtSelect ''fgSelect.Text
    
       ' cmdSelect_Click
    End If
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
fgDossier.Clear: fgDossier.Row = 0

cnADO.Open paramODBC_DSN_SAB
Exit Sub

Error_Handler:
blnControl = False

End Sub





Private Sub fgDossier_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wOrigine As String
Dim I As Integer
On Error Resume Next
If Y <= fgDossier.RowHeightMin Then
Else
    If fgDossier.Rows > 1 Then
        Call fgDossier_Color(fgDossier_RowClick, MouseMoveUsr.BackColor, fgDossier_ColorClick)
        fgDossier.Col = 0: wOrigine = Trim(fgDossier.Text)
                            
            fgDossier.Col = 1
            I = Val(fgDossier.Text)
            Select Case wOrigine
                 Case constYCREDOS0: xYCREDOS0 = meYBIACRE.YCREDOS0(I): srvYCREDOS0_ElpDisplay xYCREDOS0
                 Case constYCREPRE0: xYCREPRE0 = meYBIACRE.YCREPRE0(I): srvYCREPRE0_ElpDisplay xYCREPRE0
                 Case constYCREPLA0: xYCREPLA0 = meYBIACRE.YCREPLA0(I): srvYCREPLA0_ElpDisplay xYCREPLA0
                 Case constYCREEMP0: xYCREEMP0 = meYBIACRE.YCREEMP0(I): srvYCREEMP0_ElpDisplay xYCREEMP0
                 Case constYCREEVE0: xYCREEVE0 = meYBIACRE.YCREEVE0(I): srvYCREEVE0_ElpDisplay xYCREEVE0
                 Case constYCREAVI0: xYCREAVI0 = meYBIACRE.YCREAVI0(I): srvYCREAVI0_ElpDisplay xYCREAVI0
                 Case constYCREBIS0: xYCREBIS0 = meYBIACRE.YCREBIS0(I): srvYCREBIS0_ElpDisplay xYCREBIS0
            End Select
   End If
End If
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





Private Sub mnuPrint0_Avis_Click()
Dim K As Integer, prtI As Integer
For K = 1 To fgSelect_YCREAVI0_Nb
    For prtI = 1 To meYBIACRE.prtNb
            prtBIA_FTP_YCREAVI0 fgSelect_YCREAVI0(K), meYBIACRE
    Next prtI
Next K

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then txtSelect.SetFocus

End Sub

Private Sub SSTab1_GotFocus()
Select Case SSTab1.Tab
    Case 0: fgSelect.LeftCol = 0
End Select
End Sub

Private Sub txtSelect_Change()
Dim I As Long, X As String, lenX As Integer
On Error Resume Next
X = Trim(txtSelect)
lenX = Len(X)
fgSelect.Col = 0
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    
    If X <= Mid$(fgSelect.Text, 1, lenX) Then
        fgSelect.LeftCol = 0
        fgSelect.TopRow = I
        Exit Sub
    End If
Next I

End Sub

Private Sub txtSelect_GotFocus()
Call txt_GotFocus(txtSelect)

End Sub


Private Sub txtSelect_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)
End Sub


Private Sub txtSelect_LostFocus()
Call txt_LostFocus(txtSelect)

End Sub

Public Sub lstDossier_Load()
Dim I As Integer, K As Integer
Dim blnOk As Boolean



End Sub









