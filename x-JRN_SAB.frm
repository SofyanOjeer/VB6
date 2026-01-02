VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form X_frmJRN_SAB 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_CDO : Crédits dodumentaires"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13560
   Icon            =   "x-JRN_SAB.frx":0000
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
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Rechercher"
      TabPicture(0)   =   "x-JRN_SAB.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Courrier Ouv/Mod"
      TabPicture(1)   =   "x-JRN_SAB.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDossier"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Courrier Utilisations"
      TabPicture(2)   =   "x-JRN_SAB.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Détail fichiers"
      TabPicture(3)   =   "x-JRN_SAB.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fgDossier"
      Tab(3).ControlCount=   1
      Begin VB.Frame fraDossier 
         Height          =   8145
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   13260
         Begin VB.Frame fraDossier_Info 
            Enabled         =   0   'False
            Height          =   7695
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   7095
         End
         Begin VB.Frame fraDossier_Saisie 
            Height          =   2895
            Left            =   7320
            TabIndex        =   9
            Top             =   5160
            Width           =   5895
            Begin VB.CheckBox chkDossier_prtNb 
               Alignment       =   1  'Right Justify
               Caption         =   "Courrier en 2 exemplaires"
               Height          =   375
               Left            =   120
               TabIndex        =   13
               Top             =   1680
               Width           =   2295
            End
            Begin VB.ListBox lstDossier_Contact 
               Height          =   2400
               Left            =   3480
               TabIndex        =   10
               Top             =   240
               Width           =   2265
            End
         End
         Begin VB.ListBox lstOptions 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4650
            ItemData        =   "x-JRN_SAB.frx":037A
            Left            =   7320
            List            =   "x-JRN_SAB.frx":0381
            Style           =   1  'Checkbox
            TabIndex        =   6
            Top             =   360
            Width           =   5805
         End
      End
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   135
         TabIndex        =   3
         Top             =   330
         Width           =   13290
         Begin VB.TextBox txtSelect 
            Height          =   285
            Left            =   135
            TabIndex        =   7
            Top             =   240
            Width           =   1230
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7425
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   13080
            _ExtentX        =   23072
            _ExtentY        =   13097
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
            FormatString    =   ">Dossier       |>Prêt       |> Nature    |>Montant              |<Etat    |>Date ouverture  ||"
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
         TabIndex        =   12
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
         FormatString    =   $"x-JRN_SAB.frx":0391
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
      Picture         =   "x-JRN_SAB.frx":044B
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
End
Attribute VB_Name = "X_frmJRN_SAB"
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
Dim SAB_CRE_Aut As typeAuthorization

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim fgDossier_FormatString As String, fgDossier_K As Integer
Dim fgDossier_RowDisplay As Integer, fgDossier_RowClick As Integer, fgDossier_ColClick As Integer
Dim fgDossier_ColorClick As Long, fgDossier_ColorDisplay As Long
Dim fgDossier_Sort1 As Integer, fgDossier_Sort2 As Integer
Dim fgDossier_SortAD As Integer, fgDossier_Sort1_Old As Integer
Dim fgDossier_arrIndex As Integer
Dim blnfgDossier_DisplayLine As Boolean

Dim xElpTable As typeElpTable

Dim cnADO As New ADODB.Connection
Dim rsADO As New ADODB.Recordset
Dim meJRNENT0 As typeJRNENT0, xJRNENT0 As typeJRNENT0


Private Sub fgDossier_Display()
Dim I As Integer
On Error Resume Next
SSTab1.Tab = 1
fraDossier.Enabled = True
fgDossier_Reset

fgDossier.Rows = 1
fgDossier.FormatString = fgDossier_FormatString

For I = 0 To lstOptions.ListCount - 1
    lstOptions.Selected(I) = False
Next I


End Sub

Public Sub fgDossier_DisplayLine(lOrigine As String, lId As String, lText As String)
On Error Resume Next
fgDossier.Rows = fgDossier.Rows + 1
fgDossier.Row = fgDossier.Rows - 1
fgDossier.Col = 0: fgDossier.Text = lOrigine
fgDossier.Col = 1: fgDossier.Text = lId
fgDossier.Col = 2: fgDossier.Text = lText
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

End Sub
Public Sub fgDossier_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgDossier.Rows - 1
    fgDossier.Row = I
    fgDossier.Col = lK
    X = Format$(Val(fgDossier.Text), "0000000")
    fgDossier.Col = fgDossier_arrIndex - 1
    Select Case lK
        Case 1, 2: fgDossier.Text = X
    End Select
Next I


fgDossier_Sort1 = fgDossier_arrIndex - 1: fgDossier_Sort2 = fgDossier_arrIndex - 1
fgDossier_Sort
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
     xSQL = "select * from SAB073JRN.JRNENT0 where JOUSER = 'GRAVADE' and  JORCV = 411"
   
 '   xSQL = "select * from SAB073JRN.JRNENTLU where JOUSER = 'GUYOT' and  JODATE = '260304'"
    xSQL = "select * from SAB073JRN.JRNENT0 where   JODATE = '260304' and JOOBJ = 'ZSWIFTB0' "
 Set rsADO = cnADO.Execute(xSQL)  '$2003.11.04
    Do While Not rsADO.EOF
        'Call srvYCREPRE0_GetBuffer_ODBC(rsADO, meYCREPRE0)
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            fgSelect.Col = 0: fgSelect.Text = rsADO("JOOBJ")
           fgSelect.CellForeColor = vbBlue
            fgSelect.Col = 1: fgSelect.Text = rsADO("JORCV")
            fgSelect.Col = 2: fgSelect.Text = rsADO("JOSEQN")
   
            fgSelect.Col = 3: fgSelect.Text = rsADO("JODATE") & rsADO("JOTIME")
            fgSelect.CellForeColor = vbBlue
            
            fgSelect.Col = 4: fgSelect.Text = rsADO("JOPGM")
            fgSelect.Col = 5: fgSelect.Text = rsADO("JRNBIADOS")
  
        rsADO.MoveNext
    Loop
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

Call BiaPgmAut_Init(mId$(Msg, 1, 12), SAB_CRE_Aut)

'blnSetfocus = True
Form_Init


End Sub


Public Sub Form_Init()
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistent", vbCritical, "frmJRNENT0.param_init"
    Unload Me
Else
    lstErr.Clear
End If

blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgDossier_FormatString = fgDossier.FormatString
fgDossier.Enabled = True
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
'recJRNENT0_Init meJRNENT0
xJRNENT0 = meJRNENT0


fraDossier.Enabled = False
blnControl = True
lstDossier_Contact.ForeColor = vbMagenta

fgSelect_Display
fraDossier_Info.Enabled = False   ' La frame n'est que affichage d'informations
cmdPrint.Enabled = SAB_CRE_Aut.Saisir


End Sub


Public Function param_Init()
Dim K As Integer, K1 As Integer, X As String
Dim wText As String
Dim V

param_Init = Null
End Function






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


End Sub

Private Sub cmdSelect(lJOOBJ As String, lJORCV As Long, lJOSEQN As Long)
Dim V
Dim xSQL As String
Dim X As String

On Error GoTo Error_Handler

V = Null
Set rsADO = Nothing
X = lJOOBJ
Mid$(X, 1, 1) = "J"
xSQL = "select * from SAB073JRN." & X & " where JORCV = " & lJORCV & " And JOSEQN = " & lJOSEQN

Set rsADO = cnADO.Execute(xSQL)

Select Case mId$(lJOOBJ, 1, 4)
    Case "ZSWI": V = srvJSWI_Sql(rsADO, lJOOBJ, lJORCV, lJOSEQN)
    Case Else: V = "Non programmé"
End Select

If Not IsNull(V) Then GoTo Error_MsgBox

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & lJOOBJ & " : " & lJORCV & " : " & lJOSEQN
End Sub

Private Sub fgDossier_Click()
fgDossier.LeftCol = 0

End Sub

Private Sub fgDossier_LeaveCell()
On Error Resume Next
fgDossier.CellBackColor = &HE0E0E0
End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wJOOBJ As String, wJORCV As Long, wJOSEQN As Long
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
        fgSelect.Col = 0: wJOOBJ = Trim(fgSelect.Text)
        fgSelect.Col = 1: wJORCV = CLng(Val(fgSelect.Text))
        fgSelect.Col = 2: wJOSEQN = CLng(Val(fgSelect.Text))
        cmdSelect wJOOBJ, wJORCV, wJOSEQN

   End If
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If blnControl Then
    cnADO.Close
    Set cnADO = Nothing
End If
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

If fraDossier.Enabled Then
    fraDossier.Enabled = False
    SSTab1.Tab = 0
    fgSelect.LeftCol = 0

    Exit Sub
End If


If currentAction = "" Then
   
Else
    X = MsgBox("Voulez-vous réellement abandonner la mise à jour?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption)
    If X = vbYes Then
        currentAction = ""
    Else
        Exit Sub
    End If
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
fgDossier.Clear: fgDossier.Row = 0

cnADO.Open paramODBC_DSN_SAB 'JRN
Exit Sub

Error_Handler:
blnControl = False
MsgBox Error
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
   End If
End If
End Sub

Public Sub fgDossier_Reset()
fgDossier.Clear
fgDossier_Sort1 = 0: fgDossier_Sort2 = 0
fgDossier_Sort1_Old = -1
fgDossier_RowDisplay = 0: fgDossier_RowClick = 0
fgDossier_arrIndex = 3
blnfgDossier_DisplayLine = False
fgDossier_SortAD = 6
fgDossier.LeftCol = 0

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





Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then txtSelect.SetFocus

End Sub

Private Sub txtSelect_Change()
Dim I As Long, X As String, lenX As Integer
On Error Resume Next
X = Trim(txtSelect)
lenX = Len(X)
fgSelect.Col = 0
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    
    If X <= mId$(fgSelect.Text, 1, lenX) Then
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










