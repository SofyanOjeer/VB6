VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBIA_System 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA System"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13560
   Icon            =   "BIA_System.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9270
   ScaleWidth      =   13560
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
      TabPicture(0)   =   "BIA_System.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "........"
      TabPicture(1)   =   "BIA_System.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgSelect"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13290
         Begin VB.Frame fraSelect 
            Height          =   6855
            Left            =   120
            TabIndex        =   10
            Top             =   1200
            Width           =   13095
         End
         Begin VB.Frame fraSelect_Options 
            Height          =   1005
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   11355
            Begin VB.ComboBox cboSelect_LIB 
               Height          =   288
               Left            =   6360
               TabIndex        =   13
               Text            =   "cboSelect_LIB"
               Top             =   240
               Width           =   1452
            End
            Begin VB.ComboBox cboSelect_SQL 
               Height          =   315
               Left            =   240
               TabIndex        =   11
               Text            =   "cboSelect_SQL"
               Top             =   240
               Width           =   3615
            End
            Begin VB.TextBox txtSelect_File 
               Height          =   285
               Left            =   8160
               TabIndex        =   9
               Text            =   "C:\BIASRC\DTA\SAB073Y_X.txt"
               Top             =   240
               Width           =   3012
            End
            Begin VB.Label lblSelect_File 
               Caption         =   "Fichier"
               Height          =   252
               Left            =   5400
               TabIndex        =   8
               Top             =   240
               Width           =   720
            End
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Exécuter la requête"
            Height          =   645
            Left            =   11880
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSelect 
         Height          =   6825
         Left            =   -74880
         TabIndex        =   12
         Top             =   600
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   12039
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
         FormatString    =   "<Date        ||                           |||"
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
      Picture         =   "BIA_System.frx":047A
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
Attribute VB_Name = "frmBIA_System"
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
Dim x As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim BIA_SYSTEM_Aut As typeAuthorization
Dim curX1 As Currency, curX2 As Currency
Dim blnAuto As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim cnADO As New ADODB.Connection, rsADO As New ADODB.Recordset, errADO As ADODB.Error
Dim blnTransaction As Boolean

Dim cnSAB073Y As New ADODB.Connection, rsSAB073Y As New ADODB.Recordset

Public Sub cmdSelect_Table_Create()
'Création table suivant la syntaxe Vb 'Type typeXXXX     end Type'
'---------------------------------------------------------------
Dim xIn As String
Dim blnCreate As Boolean
Dim K As Integer
Dim xSql As String, mTable_Name As String, mTable_Fields As String
Dim wField As String, wType As String, wLen As String
On Error GoTo Error_Handler

cnSAB073Y_Open
lstErr.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "> cmdSelect_Table_Create : " & Time)
DoEvents

'C:\BIASRC\DTA\SAB073Y_X.txt
Open Trim(txtSelect_File) For Input As #1

Do Until EOF(1)
    Line Input #1, xIn
    xIn = Trim(xIn)
    If xIn = "" Or mId$(xIn, 1, 1) = "'" Then
    
    Else
'Nom de la table
'----------------
        If mId$(xIn, 1, 9) = "Type type" Then
            blnCreate = True
            mTable_Name = mId$(xIn, 10, Len(xIn) - 9)
            mTable_Fields = ""
        Else
'Création de la table
'----------------
            If xIn = "End Type" Then
                If blnCreate Then
                    blnCreate = False
                    
                    On Error Resume Next
                    xSql = "DROP TABLE " & mTable_Name
                    cnSAB073Y.Execute xSql
                    On Error GoTo Error_Handler
                    
                    Mid$(mTable_Fields, 1, 1) = " " ' supprimer la première virgule
                    xSql = "CREATE TABLE " & mTable_Name & " (" & mTable_Fields & ")"
                    Call lstErr_AddItem(lstErr, cmdContext, "> Table : " & mTable_Name)
                    Debug.Print xSql
                    cnSAB073Y.Execute xSql
                    
                    'xsql = "CREATE INDEX PrimaryKey ON " & mTable_Name & " (" & mTable_Index & ")"
                    'cnSAB073Y.Execute xSql
                End If
            Else
'Champs
'----------------
                K = 0
                wField = Trim(Space_Scan(xIn, K))      ' nom du champ
                x = Trim(Space_Scan(xIn, K))           ' as
                wType = Trim(Space_Scan(xIn, K))       ' type
                If wType = "String" Then
                    x = Trim(Space_Scan(xIn, K))      ' *
                    wLen = Trim(Space_Scan(xIn, K))
                    mTable_Fields = mTable_Fields & "," & wField & " TEXT(" & wLen & ")"
                Else
                    mTable_Fields = mTable_Fields & "," & wField & " " & wType
                End If
            End If
       End If

    End If
    
    
Loop



GoTo Exit_Sub

Error_Handler:
MsgBox Error, vbCritical, Me.Caption & " :  cmdSelect_Table_Create"
Exit_Sub:
Close
cnSAB073Y_Close
Call lstErr_AddItem(lstErr, cmdContext, "< cmdSelect_Table_Create : " & Time)

End Sub

Public Sub cmdSelect_Table_Create_DSP()
'---------------------------------------------------------------
Dim xIn As String
Dim blnCreate As Boolean
Dim K As Integer
Dim xSql As String, mTable_Name As String, mTable_Fields As String
Dim wField As String, wType As String, wLen As String
Dim xWHFLDT As String

On Error GoTo Error_Handler

cnSAB073Y_Open
lstErr.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "> cmdSelect_Table_Create : " & Time)
DoEvents

mTable_Name = UCase$(Trim(txtSelect_File))

mTable_Fields = ""

xSql = "select * from  " & paramIBM_Library_SABSPE & ".DSPFFDW0K " _
    & " where WHSYSN = '" & paramIBM_AS400_ID & "'" _
    & " and WHLIB = '" & UCase$(cboSelect_LIB) & "'" _
    & " and WHFILE = '" & mTable_Name & "'" _
    & " order by WHFOBO"
    
Set rsADO = cnADO.Execute(xSql)
Do While Not rsADO.EOF

    wField = Trim(rsADO("WHFLDI"))
    xWHFLDT = rsADO("WHFLDT")
    If xWHFLDT = "A" Then
        mTable_Fields = mTable_Fields & "," & wField & " TEXT(" & rsADO("WHFLDB") & ")"
    Else
        If xWHFLDT = "B" Then
            wType = "integer"
        Else
            Select Case rsADO("WHFLDP")
                Case 0: wType = "long"
                Case 2, 3: wType = "currency"
                Case Else: wType = "double"
            End Select
            If wField = "MBRCDC" Or wField = "MBDSZ2" Or wField = "MBDSZZ" Then wType = "currency"
        End If
        mTable_Fields = mTable_Fields & "," & wField & " " & wType
    End If
    
    rsADO.MoveNext
Loop

If mTable_Fields = "" Then
    MsgBox "champs non définis", vbCritical, "Création de la table : " & mTable_Name
    Exit Sub
End If
On Error Resume Next
xSql = "DROP TABLE " & mTable_Name
cnSAB073Y.Execute xSql
On Error GoTo Error_Handler
Mid$(mTable_Fields, 1, 1) = " " ' supprimer la première virgule
xSql = "CREATE TABLE " & mTable_Name & " (" & mTable_Fields & ")"
Call lstErr_AddItem(lstErr, cmdContext, "> Table : " & mTable_Name)
cnSAB073Y.Execute xSql
    MsgBox "Création terminée", vbOK, "Création de la table : " & mTable_Name

GoTo Exit_Sub

Error_Handler:
MsgBox Error, vbCritical, Me.Caption & " :  cmdSelect_Table_Create"
Exit_Sub:
Close
cnSAB073Y_Close
Call lstErr_AddItem(lstErr, cmdContext, "< cmdSelect_Table_Create : " & Time)

End Sub

Public Sub cnSAB073Y_Close()
On Error Resume Next

cnSAB073Y.Close
Set cnSAB073Y = Nothing


End Sub

Public Sub cnSAB073Y_Open()
On Error GoTo Error_Handler
Dim x As String

cnSAB073Y.Open paramODBC_DSN_SAB073Y

Exit Sub

Error_Handler:

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
Dim x As String, lenX As Integer
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
Dim I As Integer, x As String
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    If lK = 2 Then
        fgSelect.Col = 2
        x = fgSelect.Text
    Else
        x = ""
    End If
    
    fgSelect.Col = 3
    x = x & Format$(Val(fgSelect.Text), "000000000000000.00")
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = x
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
'For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), BIA_SYSTEM_Aut)
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
    If Not blnAuto Then MsgBox "paramétrage inconsistant", vbCritical, "frmBIA_SYSTEM.paramSAA_Init"
    Unload Me
Else
    lstErr.Clear
End If


blnControl = False
fgSelect_FormatString = fgSelect.FormatString

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
Call lstErr_Clear(lstErr, cmdContext, "BIA_SYSTEM : param_init"): DoEvents

fraSelect.Visible = False


cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1 - Création Table => SAB073Y"
cboSelect_SQL.AddItem "2 - "
cboSelect_SQL.ListIndex = 0

cboSelect_LIB.Clear
cboSelect_LIB.AddItem "SAB073"
cboSelect_LIB.AddItem "SAB073SPE"
cboSelect_LIB.AddItem "SAB073JRN"
cboSelect_LIB.AddItem "JPLTST"
cboSelect_LIB.ListIndex = 0

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



Private Sub cboSelect_LIB_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
Dim x As String
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL"

Select Case mId$(cboSelect_SQL.Text, 1, 1)
    Case 1:  cmdSelect_Table_Create '_DSP
End Select

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub
Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, NB As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_SYSTEM_cmdSelect_Ok ........"): DoEvents

'fgSelect.Clear
If blnOk Then
    cmdSelect_Ok.Caption = "Options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = False
    cmdSelect_SQL
Else
    cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = True
    fraSelect.Visible = False
End If
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_SYSTEM_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0


End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim K As Long
On Error Resume Next
If Y <= fgSelect.RowHeightMin Then
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





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
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


Private Sub mnuSelect_Print_Liste_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
'cmdPrint_List1_Ok
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

cnADO.Close
Set cnADO = Nothing

End Sub

Public Sub cnADO_Open()
On Error GoTo Error_Handler
Dim x As String
cnADO.Open paramODBC_DSN_SAB


Exit Sub

Error_Handler:
blnControl = False
If Not blnAuto Then MsgBox Error

End Sub


Private Sub txtSelect_File_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


