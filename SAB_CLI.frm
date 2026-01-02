VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSAB_CLI 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_CLI : Clients"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   Icon            =   "SAB_CLI.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6210
   ScaleWidth      =   9090
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   5100
      TabIndex        =   4
      Top             =   -15
      Width           =   3495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   30
      TabIndex        =   2
      Top             =   525
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Rechercher"
      TabPicture(0)   =   "SAB_CLI.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Import"
      TabPicture(1)   =   "SAB_CLI.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTAb1"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraTAb1 
         Height          =   4905
         Left            =   -74970
         TabIndex        =   7
         Top             =   525
         Width           =   8940
      End
      Begin VB.Frame fraTab0 
         Height          =   4875
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   8850
         Begin VB.Frame fraSelect 
            Height          =   705
            Left            =   45
            TabIndex        =   5
            Top             =   180
            Width           =   8760
            Begin VB.CommandButton cmdExport 
               Caption         =   "Exporter adresses banques"
               Height          =   435
               Left            =   1320
               TabIndex        =   9
               Top             =   240
               Width           =   2580
            End
            Begin VB.CommandButton cmdSelect_BANQ 
               Caption         =   "sélectionner les banques"
               Height          =   435
               Left            =   6000
               TabIndex        =   8
               Top             =   120
               Width           =   2580
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   3870
            Left            =   120
            TabIndex        =   6
            Top             =   915
            Width           =   8730
            _ExtentX        =   15399
            _ExtentY        =   6826
            _Version        =   393216
            Rows            =   1
            Cols            =   7
            FixedCols       =   0
            RowHeightMin    =   200
            BackColor       =   14737632
            ForeColor       =   12582912
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
            FormatString    =   $"SAB_CLI.frx":0342
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
      Left            =   8600
      Picture         =   "SAB_CLI.frx":040F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint_Adresse 
         Caption         =   "Imprimer Adresses"
      End
      Begin VB.Menu mnuPrint_BIC 
         Caption         =   "Imprimer liste client / BIC"
      End
      Begin VB.Menu mnuX1 
         Caption         =   "-"
      End
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
Attribute VB_Name = "frmSAB_CLI"
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
Dim SAB_CLI_Aut As typeAuthorization


Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim meYCLIENA0 As typeYCLIENA0, xYCLIENA0 As typeYCLIENA0
Dim meYADRESS0 As typeYADRESS0, xYADRESS0 As typeYADRESS0
Dim blnError As Boolean
Dim meMVTP0 As typeMvtP0, xMvtP0 As typeMvtP0, wMVTP0 As typeMvtP0
Dim meYCLIREF0 As typeYCLIREF0, xYCLIREF0 As typeYCLIREF0

Dim xElpKMInfo As typeElpKMInfo
Dim xElpKMIndex As typeElpKMIndex

Dim cnADO As New ADODB.Connection
Dim rsADO As New ADODB.Recordset

Private Sub fgSelect_Display()

SSTab1.Tab = 0

fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString


End Sub

Public Sub fgSelect_DisplayLine(lOrigine As String, lId As String, lText As String)
On Error Resume Next
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1
fgSelect.Col = 0: fgSelect.Text = lOrigine
fgSelect.Col = 1: fgSelect.Text = lId
fgSelect.Col = 2: fgSelect.Text = lText
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
    X = Format$(Val(fgSelect.Text), "0000000")
    fgSelect.Col = fgSelect_arrIndex - 1
    Select Case lK
        Case 1, 2: fgSelect.Text = X
    End Select
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

Call BiaPgmAut_Init(mId$(Msg, 1, 12), SAB_CLI_Aut)

'blnSetfocus = True
Form_Init


End Sub


Public Sub Form_Init()
Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistent", vbCritical, "frmYCLIENA0.param_init"
    Unload Me
End If

blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fraTAb1.Visible = SAB_CLI_Aut.Xspécial
tableMvtP0_Open
cmdReset
Me.Enabled = True

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
recYCLIENA0_Init meYCLIENA0
xYCLIENA0 = meYCLIENA0

paramYCLIENA0_Import = paramTemp_Folder & "FTP\YCLIENA0.txt"
paramYCLIREF0_Import = paramTemp_Folder & "FTP\YCLIREF0.txt"
paramYADRESS0_Import = paramTemp_Folder & "FTP\YADRESS0.txt"

blnControl = True

End Sub



Public Function param_Init()
Dim K As Integer, K1 As Integer, X As String

Dim V
param_Init = Null


End Function






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

Private Sub cmdSelect_CLIENAETA(lValue As String)
Dim blnOk As Boolean
Dim xsql As String

On Error GoTo Error_Handler

fgSelect_Display

fgSelect.Visible = False


    Set rsADO = Nothing
    
    xsql = "select * from ZCLIENA0 WHERE CLIENAETA LIKE 'BA%'"
    Set rsADO = cnADO.Execute(xsql)  '$2003.11.04

    Do While Not rsADO.EOF
        'Call srvYCREPRE0_GetBuffer_ODBC(rsADO, meYCREPRE0)
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            fgSelect.Col = 0: fgSelect.Text = rsADO("CLIENACLI")
            fgSelect.CellForeColor = vbBlue
            fgSelect.Col = 1: fgSelect.Text = rsADO("CLIENASIG")
            fgSelect.Col = 2: fgSelect.Text = rsADO("CLIENARA1")
  
        rsADO.MoveNext
    Loop
fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort

fgSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdExport_Exe()
Dim K As Integer, xsql As String
Dim xFile As String, X As String
On Error GoTo Error_Handler

xFile = "C:\Temp\BIA_BANQUE.txt"
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> cmdExport_ ........"): DoEvents
Open xFile For Output As #2

For I = 1 To fgSelect.Rows - 1
    
    fgSelect.Row = I
    fgSelect.Col = 0
    
    xsql = "select * from ZCLIENA0 where CLIENACLI = '" & Trim(fgSelect.Text) & "'"
    Set rsADO = cnADO.Execute(xsql)
    If Not rsADO.EOF Then
        V = srvYCLIENA0_GetBuffer_ODBC(rsADO, xYCLIENA0)
        If Not IsNull(V) Then
            MsgBox V, vbCritical, "prtSAB_CLI_Adresse : Lecture ZCLIENT0 : "
            Exit For
        End If
' recherche ADRESSE
'-------------------
        xYADRESS0.ADRESSNUM = xYCLIENA0.CLIENACLI
        Call srvYADRESS0_Client(xYADRESS0, cnADO)
        Print #2, xYCLIENA0.CLIENACLI & ";" _
        & Trim(xYADRESS0.ADRESSRA1) & " " & Trim(xYADRESS0.ADRESSRA2) & ";" _
        & Trim(xYADRESS0.ADRESSAD1) & ";" _
         & Trim(xYADRESS0.ADRESSAD2) & ";" _
        & Trim(xYADRESS0.ADRESSCOP) & " " & Trim(xYADRESS0.ADRESSVIL) & ";" _
        & Trim(xYADRESS0.ADRESSPAY)
     End If
Next I
Close #2


Call lstErr_AddItem(lstErr, cmdContext, "< cmdExport : " & fgSelect.Rows - 1): DoEvents
GoTo Exit_Sub

Error_Handler:
    Close
    MsgBox Error, vbCritical, Me.Caption & " :  " & xFile
Exit_Sub:
    Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdExport_Click()
cmdExport_Exe
End Sub

Private Sub cmdPrint_Click()
Select Case SSTab1.Tab
    Case 0:
            If fgSelect.Rows > 1 Then
                Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
           End If
End Select

End Sub

Private Sub cmdSelect_BANQ_Click()
cmdSelect_CLIENAETA "BA"
End Sub

Private Sub fgSelect_Click()
fgSelect.LeftCol = 0

End Sub

Private Sub fgSelect_LeaveCell()
On Error Resume Next
fgSelect.CellBackColor = &HE0E0E0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cnADO.Close
    Set cnADO = Nothing

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
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False
fgSelect.Clear: fgSelect.Row = 0
cnADO.Open paramODBC_DSN_SAB




End Sub





Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wOrigine As String
On Error Resume Next
If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = 0: wOrigine = fgSelect.Text
        fgSelect.Col = 2
        MsgTxt = Space$(34) & fgSelect.Text
        MsgTxtIndex = 0
                            
        Select Case wOrigine
            Case constYCLIENA0: srvYCLIENA0_GetBuffer xYCLIENA0: srvYCLIENA0_ElpDisplay xYCLIENA0
            Case constYCLIREF0: srvYCLIREF0_GetBuffer xYCLIREF0: srvYCLIREF0_ElpDisplay xYCLIREF0
            Case constYADRESS0: srvYADRESS0_GetBuffer xYADRESS0: srvYADRESS0_ElpDisplay xYADRESS0
       End Select
        
   End If
End If
End Sub

Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 3
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = 0

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

End Sub




Private Sub mnuPrint_Adresse_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

prtSAB_CLI_Adresse fgSelect, cnADO
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuPrint_BIC_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
prtSAB_CLI_BIC fgSelect, cnADO
Me.Enabled = True: Me.MousePointer = 0

End Sub


