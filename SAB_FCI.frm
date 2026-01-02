VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSAB_FCI 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_FCI : courrier"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "SAB_FCI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
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
      TabCaption(0)   =   "Sélection des courriers "
      TabPicture(0)   =   "SAB_FCI.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "?"
      TabPicture(1)   =   "SAB_FCI.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraSelect 
         Height          =   8445
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13560
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   645
            Left            =   11880
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   360
            Width           =   1095
         End
         Begin VB.Frame fraSelect_Options 
            Height          =   1005
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   11355
            Begin VB.TextBox txtSelect_FCIGCOCPT 
               Height          =   285
               Left            =   5880
               TabIndex        =   12
               Top             =   360
               Width           =   1815
            End
            Begin VB.CheckBox chkSelect_FCIGCODAJ 
               Caption         =   "Courriers du"
               Height          =   255
               Left            =   360
               TabIndex        =   7
               Top             =   360
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker txtSelect_AMJMIN 
               Height          =   300
               Left            =   1920
               TabIndex        =   8
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
               Format          =   121700355
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtSelect_AMJMAX 
               Height          =   300
               Left            =   3360
               TabIndex        =   9
               Top             =   600
               Visible         =   0   'False
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
               Format          =   121700355
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblSelect_FCIGCOCPT 
               Caption         =   "Compte"
               Height          =   255
               Left            =   4680
               TabIndex        =   10
               Top             =   360
               Width           =   600
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect_D 
            Height          =   7095
            Left            =   5280
            TabIndex        =   13
            Top             =   1320
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   12515
            _Version        =   393216
            Cols            =   4
            BackColorFixed  =   8438015
            BackColorBkg    =   12640511
            GridColor       =   255
            WordWrap        =   -1  'True
            GridLines       =   3
            AllowUserResizing=   3
            FormatString    =   "Champ           |Intitulé               |<Valeur                                                                                |"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6825
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   13200
            _ExtentX        =   23283
            _ExtentY        =   12039
            _Version        =   393216
            Rows            =   1
            Cols            =   14
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
            FormatString    =   $"SAB_FCI.frx":0044
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
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "SAB_FCI.frx":00F3
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
Attribute VB_Name = "frmSAB_FCI"
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
Dim SAB_FCIAut As typeAuthorization
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

'______________________________________________________________________

Dim xZFCIGCO0 As typeZFCIGCO0
Dim arrZFCIGCO0() As typeZFCIGCO0, arrZFCIGCO0_Nb As Long, arrZFCIGCO0_Max As Long, arrZFCIGCO0_Index As Long

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
Private Sub fgSelect_D_Display()
Dim V
Dim xSQL As String
On Error GoTo Error_Handler

V = Null
fgSelect_D.Clear
fgSelect_D.Visible = False
fgSelect_D.FormatString = fgSelect_D_FormatString

Call srvZFCIGCO0_fgDisplay(xZFCIGCO0, fgSelect_D)
fgSelect_D.Visible = True
cmdPrint.Enabled = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : fgSelect_D_Display"
End Sub


Private Sub fgSelect_Display()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
cmdPrint.Enabled = False

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgselect_Display"
    
For I = 1 To arrZFCIGCO0_Nb
         
    xZFCIGCO0 = arrZFCIGCO0(I)
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
Next I

fgSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Courriers : " & arrZFCIGCO0_Nb): DoEvents
If fgSelect.Rows > 1 Then
    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
    cmdPrint.Enabled = True
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
On Error Resume Next
fgSelect.Col = 0:
If Trim(xZFCIGCO0.FCIGCOCPT) <> "" Then
    fgSelect.Text = xZFCIGCO0.FCIGCOCPT
Else
    fgSelect.Text = xZFCIGCO0.FCIGCONDE
End If
fgSelect.Col = 1: fgSelect.Text = xZFCIGCO0.FCIGCONUC
fgSelect.Col = 2: fgSelect.Text = xZFCIGCO0.FCIGCODAJ
fgSelect.Col = 5: fgSelect.Text = xZFCIGCO0.FCIGCOCOU
fgSelect.Col = 6: fgSelect.Text = xZFCIGCO0.FCIGCOLIB
fgSelect.Col = 3: fgSelect.Text = xZFCIGCO0.FCIGCOREJ
fgSelect.Col = 4: fgSelect.Text = xZFCIGCO0.FCIGCOLIR

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

End Sub



Public Sub cmdSendMail()
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim wPath As String
Dim xDétail As String
Dim wNb As Long, xNb As String
Dim iRow As Integer, K As Integer
Dim X As String

wNb = 0
xDétail = ""
'srvZFCIGCO0_Init xZFCIGCO0

For iRow = 1 To fgSelect.Rows - 1
    
    fgSelect.Row = iRow
    fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
    xZFCIGCO0 = arrZFCIGCO0(K)
        wNb = wNb + 1
        If Trim(xZFCIGCO0.FCIGCOCPT) <> "" Then
             xDétail = xDétail & "<TR><TD width = 40%>" & htmlFontColor("RED") & "<FONT SIZE=-1>" & xZFCIGCO0.FCIGCOLIB & "</FONT SIZE></TD>" _
                              & "<TD width = 20%><B>" & htmlFontColor("RED") & xZFCIGCO0.FCIGCOCPT & "</B></TD>" _
                              & "<TD width = 40%><FONT SIZE=-1>" & htmlFontColor("RED") & xZFCIGCO0.FCIGCONOT & "</FONT></TD></TR>"
             X = ZCHQCOM0_Sql(xZFCIGCO0.FCIGCOCPT)
             xDétail = xDétail & "<TR><TD width = 40%>" & htmlFontColor("MAGENTA") & "<FONT SIZE=-1>" & "gestion compte/carnet" & "</FONT SIZE></TD>" _
                              & "<TD width = 20%>" & htmlFontColor("MAGENTA") & xZFCIGCO0.FCIGCOCPT & "</TD>" _
                              & "<TD width = 40%><FONT SIZE=-1>" & htmlFontColor("MAGENTA") & X & "</FONT></TD></TR>"
             X = ZCHQHIS0_Sql(xZFCIGCO0.FCIGCOCPT)
             xDétail = xDétail & "<TR><TD width = 40%>" & htmlFontColor("MAGENTA") & "<FONT SIZE=-1>" & "historique des chéquiers" & "</FONT SIZE></TD>" _
                              & "<TD width = 20%>" & htmlFontColor("MAGENTA") & xZFCIGCO0.FCIGCOCPT & "</TD>" _
                              & "<TD width = 40%><FONT SIZE=-1>" & htmlFontColor("MAGENTA") & X & "</FONT></TD></TR>"
             X = ZCHQDEM0_Sql(xZFCIGCO0.FCIGCOCPT)
             xDétail = xDétail & "<TR><TD width = 40%>" & htmlFontColor("MAGENTA") & "<FONT SIZE=-1>" & "Demande de chéquiers" & "</FONT SIZE></TD>" _
                              & "<TD width = 20%>" & htmlFontColor("MAGENTA") & xZFCIGCO0.FCIGCOCPT & "</TD>" _
                              & "<TD width = 40%><FONT SIZE=-1>" & htmlFontColor("MAGENTA") & X & "</FONT></TD></TR>"
        Else
           xDétail = xDétail & "<TR><TD width = 40%>" & htmlFontColor("BLUE") & "<FONT SIZE=-1>" & xZFCIGCO0.FCIGCOLIB & "</FONT SIZE></TD>" _
                              & "<TD width = 20%><B>" & htmlFontColor("BLUE") & xZFCIGCO0.FCIGCONDE & "</B></TD>" _
                              & "<TD width = 40%><FONT SIZE=-1>" & htmlFontColor("BLUE") & xZFCIGCO0.FCIGCONRD & "</FONT></TD></TR>"
        End If
    

Next iRow

If wNb = 0 Then
    xNb = "aucun courrier FCI généré le " & dateImp(YBIATAB0_DATE_CPT_J)
Else
    xNb = " " & wNb & " courriers FCI générés le " & dateImp(YBIATAB0_DATE_CPT_J) & "<BR>"
End If
If xDétail <> "" Then
    xNb = xNb & "<BR><U></CENTER>Courriers / Comptes  :</B></U>" _
    & "<TABLE width= 80% border=1>"
End If
    
'  "<align=" & Asc34 & "left" & Asc34 & "
wSendMail.FromDisplayName = "FCI_COURRIER"
wSendMail.RecipientDisplayName = "FCI"

bgColor = "CYAN"
wSendMail.Subject = "FCI courriers du " & dateImp(YBIATAB0_DATE_CPT_J)
wSendMail.Attachment = ""
wSendMail.Message = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">" _
                    & "<FONT face=" & Asc34 & prtFontName_Arial & Asc34 & ">" _
                    & htmlFontColor("BLUE") & "<B><CENTER>" & xNb _
                    & "<BR>" & xDétail

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail


'                X = "<body bgcolor=" & Asc34 & "BLUE" & Asc34 & ">" _
'                        & "<FONT face=" & Asc34 & prtFontName_Comic & Asc34 & ">" _
'                        & htmlFontColor("WHITE") & "<BR><BR>" & fgSelect.Rows - 1 & " déclaration(s) de chèques impayés le " & dateImp(YBIATAB0_DATE_CPT_J)
'
'                Call Email_Alerte("FCI_COURRIER", "FCI", "FCI courriers du " & dateImp(YBIATAB0_DATE_CPT_J), X, True, "")

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
blnControl = False
If fgSelect_D.Visible Then fgSelect_D.Visible = False: Exit Sub
If currentAction <> "" Then
    X = MsgBox("Voulez-vous réellement abandonner la mise à jour?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
    If X = vbYes Then
        currentAction = ""
    Else
        Exit Sub
    End If
End If


lstErr.Clear
If SSTab1.Tab > 0 Then
    SSTab1.Tab = SSTab1.Tab - 1
Else
End If
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

libRéférenceInterne = ""
Call DTPicker_Set(txtSelect_AMJMIN, DSys)
Call DTPicker_Set(txtSelect_AMJMAX, YBIATAB0_DATE_CPT_J)
cmdSelect_Ok_Click
blnControl = True
End Sub
Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

SSTab1.Tab = 0

blnControl = False
fgSelect_D.Visible = False

fgSelect_FormatString = fgSelect.FormatString
fgSelect_D_FormatString = fgSelect_D.FormatString

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
If fgSelect_D.Visible Then
    prtSAB_FCI_Monitor xZFCIGCO0
Else
    Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
End If
Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long


blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAB_FCI_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
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
End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_FCI_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0


End Sub


Private Sub cmdSelect_SQL()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim xDate10 As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL"
xWhere = ""
X = Trim(txtSelect_FCIGCOCPT)
If chkSelect_FCIGCODAJ = "1" Then
    Call DTPicker_Control(txtSelect_AMJMIN, wAMJMin)
    xDate10 = Format$(Mid$(wAMJMin, 7, 2) & Mid$(wAMJMin, 5, 2) & Mid$(wAMJMin, 1, 4), "@@-@@-@@@@")

'    Call DTPicker_Control(txtSelect_AMJMAX, wAmjMax)
'     wAmjMax7 = dateIBM(wAmjMin)
   xWhere = " where FCIGCODAJ = '" & xDate10 & "'"
    If X <> "" Then xWhere = xWhere & " and FCIGCOCPT like '" & X & "%'"
Else
    If X <> "" Then xWhere = "where FCIGCOCPT like '" & X & "%'"
End If
arrZFCIGCO0_SQL xWhere


fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub arrZFCIGCO0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrZFCIGCO0(101)
arrZFCIGCO0_Max = 100: arrZFCIGCO0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SAB & ".ZFCIGCO0 " & xWhere & " order by FCIGCOCPT,FCIGCODAJ"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsZFCIGCO0_GetBuffer(rsSab, xZFCIGCO0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgselect_Display"
        '' Exit Sub
     Else
         arrZFCIGCO0_Nb = arrZFCIGCO0_Nb + 1
         If arrZFCIGCO0_Nb > arrZFCIGCO0_Max Then
             arrZFCIGCO0_Max = arrZFCIGCO0_Max + 50
             ReDim Preserve arrZFCIGCO0(arrZFCIGCO0_Max)
         End If
         
         arrZFCIGCO0(arrZFCIGCO0_Nb) = xZFCIGCO0
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
        xZFCIGCO0 = arrZFCIGCO0(K)
        fgSelect_D_Display
        
   End If
End If
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

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), SAB_FCIAut)

blnSetfocus = True
Form_Init

Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case "@AUTO_FCI": blnAuto = True
                If paramEnvironnement = constProduction Then
                    meUnit.Id = "SOBF"
                    Table_Unit meUnit
                    Printer_Set meUnit.Printer
                End If
                fraSelect_Options.Enabled = True
                Call DTPicker_Set(txtSelect_AMJMIN, YBIATAB0_DATE_CPT_J)
                cmdSelect_Ok_Click
               If fgSelect.Rows > 1 Then mnuPrint0_All_Click
                
                cmdSendMail

                Unload Me

    Case Else: blnAuto = False
End Select


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
    
For I = 1 To arrZFCIGCO0_Nb
    fgSelect.Row = I
    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
    xZFCIGCO0 = arrZFCIGCO0(K)
    prtSAB_FCI_Monitor xZFCIGCO0
Next I

Me.Show

Me.Enabled = True: Me.MousePointer = 0



End Sub



