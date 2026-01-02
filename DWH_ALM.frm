VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDWH_ALM 
   AutoRedraw      =   -1  'True
   Caption         =   "ALM : exportation .xlsx"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "DWH_ALM.frx":0000
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
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "DWH_ALM.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSource"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "réservé informatique"
      TabPicture(1)   =   "DWH_ALM.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraSource 
         Height          =   8445
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   13560
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   555
            Left            =   11340
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   285
            Width           =   1335
         End
         Begin VB.Frame fraSelect_Options 
            BackColor       =   &H00F0FFFF&
            Height          =   1005
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   10995
            Begin MSComCtl2.DTPicker txtSelect_DGAPPISPER 
               Height          =   300
               Left            =   480
               TabIndex        =   7
               Top             =   500
               Width           =   1332
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   56229891
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_DGAPPISPER 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Période"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   8
               Top             =   200
               Width           =   612
            End
         End
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "DWH_ALM.frx":0044
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
      Begin VB.Menu mnuselecté 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmDWH_ALM"
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
Dim SAB_MNUAut As typeAuthorization
Dim blnTransaction As Boolean


Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnSetfocus As Boolean

Dim blnAuto As Boolean, blnAuto_Ok As Boolean
Dim wAmjMax As String

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim wFile As String, wFilex As String

Dim almT1(100) As typeALM, blnT1 As Boolean
Dim almT2(100) As typeALM, blnT2 As Boolean
Dim almT3(100) As typeALM, blnT3 As Boolean
Dim almT4(100) As typeALM, blnT4 As Boolean
Dim almT5(100) As typeALM, blnT5 As Boolean

Dim arrDGAPPISVEC(8) As String, arrDGAPPISVEC_AMJ(8) As Long

Dim meCV1 As typeCV, meCV2 As typeCV

Dim xDGAPPIS0 As typeDGAPPIS0
'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
blnControl = False
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
lstUsr.BackColor = &HE0E0E0

cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
currentAction = ""

blnAuto = False
blnAuto_Ok = False

blnControl = True
End Sub
Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

SSTab1.Tab = 0
wAmjMax = dateFinDeMois(YBIATAB0_DATE_CPT_J)
If wAmjMax > YBIATAB0_DATE_CPT_J Then wAmjMax = dateElp("FinDeMoisP", 0, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtSelect_DGAPPISPER, wAmjMax) '


blnControl = False

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


Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_DWH_ALM_cmdSelect_Ok ........"): DoEvents
    
    cmdSelect_SQL
    
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_DWH_ALM_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_SQL()
Dim V, K As Integer, K2 As Integer, Kech As Integer
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean
Dim wGAPPISTPR As String, wGAPPISRUB As String, wGAPPISNAT As String, wGAPPISOPE As String
Dim blnExclure As Boolean, blnExclure_DBQ As Boolean, blnPassif As Boolean
Dim mGAPPISNUO As String, mGAPPISNUO_MON As Currency
Dim wT_Row As Integer, mDetail_Row As Long, mGAPPISMON As Currency
On Error GoTo Error_Handler

currentAction = "DWH_ALM : cmdSelect_SQL"
blnOk = False
   
   
Call DTPicker_Control(txtSelect_DGAPPISPER, wAmjMax)
xWhere = " where DGAPPISPER = " & wAmjMax

meCV1.OpéAmj = wAmjMax


wFile = "C:\temp\ALM_DGAPPIS0.xlsx"
'______________________________________________
If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "DWH_ALM  : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If
If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile
'_________________________________________

'$JPL 2015-11-16
'mGAPPISNUO = InputBox("Liste des obligations à reporter du tableau 1 vers le tableau 2: " _
'    & vbCrLf & " exemple 21652 21694 ")
mGAPPISNUO = ""

mGAPPISNUO_MON = 0

cmdSelect_DWH_ALM_Init

Set wsExcel = wbExcel.Sheets(6)
mDetail_Row = 1


xSql = "select * from " & paramIBM_Library_BODWH & ".DGAPPIS0 " & xWhere & " and GAPPISTAB = 1 order by GAPPISTPR"
Set rsSab = cnsab.Execute(xSql)


Do While Not rsSab.EOF
    V = rsDGAPPIS0_GetBuffer(rsSab, xDGAPPIS0)
    mGAPPISMON = xDGAPPIS0.GAPPISMON
    If xDGAPPIS0.GAPPISDEV <> "EUR" Then
        meCV1.DeviseIso = xDGAPPIS0.GAPPISDEV
        
        meCV1.Montant = xDGAPPIS0.GAPPISMON
        Call CV_Calc("", meCV1, meCV2)
        xDGAPPIS0.GAPPISMON = meCV2.Montant
        
        If xDGAPPIS0.GAPPISTP1 <> 0 Then
             meCV1.Montant = xDGAPPIS0.GAPPISTP1
            Call CV_Calc("", meCV1, meCV2)
            xDGAPPIS0.GAPPISTP1 = meCV2.Montant
        End If
        
        If xDGAPPIS0.GAPPISTP2 <> 0 Then
             meCV1.Montant = xDGAPPIS0.GAPPISTP2
            Call CV_Calc("", meCV1, meCV2)
            xDGAPPIS0.GAPPISTP2 = meCV2.Montant
        End If
        
        If xDGAPPIS0.GAPPISTM1 <> 0 Then
             meCV1.Montant = xDGAPPIS0.GAPPISTM1
            Call CV_Calc("", meCV1, meCV2)
            xDGAPPIS0.GAPPISTM1 = meCV2.Montant
        End If
        
        If xDGAPPIS0.GAPPISTM2 <> 0 Then
             meCV1.Montant = xDGAPPIS0.GAPPISTM2
            Call CV_Calc("", meCV1, meCV2)
            xDGAPPIS0.GAPPISTM2 = meCV2.Montant
        End If
   End If
    
    If xDGAPPIS0.GAPPISSEN = "A" Then xDGAPPIS0.GAPPISMON = -xDGAPPIS0.GAPPISMON
    wGAPPISTPR = Trim(xDGAPPIS0.GAPPISTPR)
    wGAPPISRUB = Trim(xDGAPPIS0.GAPPISRUB)
    wGAPPISNAT = Trim(xDGAPPIS0.GAPPISNAT)
    wGAPPISOPE = Trim(xDGAPPIS0.GAPPISOPE)
    
'$JPL 2015-11-16
    If wGAPPISOPE = "PLA" Then
        If wGAPPISTPR <> "TIT060" Then xDGAPPIS0.GAPPISTVF = "F"
    End If
    
    'If xDGAPPIS0.DGAPPISSEQ = 5215 Then
    '    Debug.Print 5215
    'End If
    
    blnExclure = False
    blnExclure_DBQ = False
'==========================================================================================================
    blnT1 = False
    blnPassif = False
    wT_Row = 0
    Select Case wGAPPISRUB
        Case "303924", "303131"
        
        Case "361611", "365621"
                    blnT1 = True: wT_Row = 36
                    almT1(36).Mt1 = almT1(36).Mt1 + xDGAPPIS0.GAPPISMON
        Case Else
            Select Case wGAPPISTPR
                Case "IMMO"
                            blnT1 = True: wT_Row = 31
                            almT1(31).Mt1 = almT1(31).Mt1 + xDGAPPIS0.GAPPISMON
                Case "CCRDTX", "CCX048", "DTX030", "DTX048", "DTX060", "PRX048", "PRODTX"
                            blnT1 = True: wT_Row = 32
                            almT1(32).Mt1 = almT1(32).Mt1 + xDGAPPIS0.GAPPISMON
                Case "PRCPAR", "TITPAR"
                            blnT1 = True: wT_Row = 33
                            almT1(33).Mt1 = almT1(33).Mt1 + xDGAPPIS0.GAPPISMON
                Case "TIT070", "PRO070"
                                blnExclure = True
                                blnT1 = True: wT_Row = 34
                                almT1(34).Mt1 = almT1(34).Mt1 + xDGAPPIS0.GAPPISMON
                Case "TIT060", "PRO060"
                            If InStr(mGAPPISNUO, xDGAPPIS0.GAPPISNUO) = 0 Then
                                blnT1 = True: wT_Row = 35
                                almT1(35).Mt1 = almT1(35).Mt1 + xDGAPPIS0.GAPPISMON
                             End If
                            
                Case "PROVIS"
                                blnT1 = True: wT_Row = 36
                                almT1(36).Mt1 = almT1(36).Mt1 + xDGAPPIS0.GAPPISMON
                Case "CAPITA", "RESLEG", "REPORT": blnPassif = True
                            blnT1 = True: wT_Row = 41
                            almT1(41).Mt1 = almT1(41).Mt1 + xDGAPPIS0.GAPPISMON
'$JPL 2015-11-16
                'Case "RISPAY": blnPassif = True
                '            blnT1 = True
                '            almT1(42).Mt1 = almT1(42).Mt1 + xDGAPPIS0.GAPPISMON
            End Select
    End Select


    Select Case xDGAPPIS0.DGAPPISVEC
        Case "01M": K2 = 0: Kech = 11
        Case "03M": K2 = 1: Kech = 12
        Case "06M": K2 = 2: Kech = 13
        Case "01A": K2 = 3: Kech = 14
        Case "02A": K2 = 4: Kech = 15
        Case Else: K2 = 5: Kech = 16
    End Select
'_________________________________________________________________________________________________________
    If blnT1 Then
        For K = 1 To K2
            If blnPassif Then
                almT1(K).Mt2 = almT1(K).Mt2 + xDGAPPIS0.GAPPISMON
            Else
                almT1(K).Mt1 = almT1(K).Mt1 + xDGAPPIS0.GAPPISMON
            End If
        Next K
        
        If xDGAPPIS0.GAPPISTVF = "V" Then
            If blnPassif Then
                almT1(Kech).Mt4 = almT1(Kech).Mt4 + xDGAPPIS0.GAPPISMON
            Else
                almT1(Kech).Mt2 = almT1(Kech).Mt2 + xDGAPPIS0.GAPPISMON
            End If
        Else
            If blnPassif Then
                almT1(Kech).Mt3 = almT1(Kech).Mt3 + xDGAPPIS0.GAPPISMON
            Else
                almT1(Kech).Mt1 = almT1(Kech).Mt1 + xDGAPPIS0.GAPPISMON
            End If
        End If
        
        Kech = Kech + 10
        almT1(Kech).Mt1 = almT1(Kech).Mt1 + xDGAPPIS0.GAPPISTP1
        almT1(Kech).Mt2 = almT1(Kech).Mt2 + xDGAPPIS0.GAPPISTP2
        almT1(Kech).Mt3 = almT1(Kech).Mt3 + xDGAPPIS0.GAPPISTM1
        almT1(Kech).Mt4 = almT1(Kech).Mt4 + xDGAPPIS0.GAPPISTM2
        almT1(Kech).Mt5 = almT1(Kech).Mt5 + xDGAPPIS0.GAPPISMAR
        '=================
        GoTo DGAPPIS0_Next
        '=================
      
    End If
'==========================================================================================================
    blnT2 = False
    blnPassif = False
    
    Select Case wGAPPISRUB
        Case "303131"
'$JPL 2015-11-16
'                    blnExclure = True
'                    blnT2 = True
'                    almT2(32).Mt1 = almT2(32).Mt1 + xDGAPPIS0.GAPPISMON
'$JPL 2015-11-16
'        Case "303924"
'                    blnT1 = True
'                    almT1(34).Mt1 = almT1(34).Mt1 + xDGAPPIS0.GAPPISMON
        Case Else
'            Select Case wGAPPISTPR
'                Case "TIT060", "PRO060"
'                            If InStr(mGAPPISNUO, xDGAPPIS0.GAPPISNUO) > 0 Then
'                                blnT1 = True
'                                almT2(34).Mt1 = almT2(34).Mt1 + xDGAPPIS0.GAPPISMON
'                            End If
'            End Select
            Select Case wGAPPISNAT
'                Case "OCT"
'                            blnT2 = True
'                            almT2(33).Mt1 = almT2(33).Mt1 + xDGAPPIS0.GAPPISMON
                Case "GEN", "BDF": blnPassif = True
                        If xDGAPPIS0.DGAPPISCLI = 11001 Then
                            blnT2 = True: wT_Row = 41
                            blnExclure_DBQ = True
                            almT2(41).Mt1 = almT2(41).Mt1 + xDGAPPIS0.GAPPISMON
                        Else
                            If xDGAPPIS0.DGAPPISCLI = 11012 Then
                                blnT2 = True: wT_Row = 42
                                blnExclure_DBQ = True
                                almT2(42).Mt1 = almT2(42).Mt1 + xDGAPPIS0.GAPPISMON
                            End If
                        End If
                        
            End Select
    End Select
    

    Select Case xDGAPPIS0.DGAPPISVEC
        Case "01M": K2 = 0: Kech = 11
        Case "03M": K2 = 1: Kech = 12
        Case "06M": K2 = 2: Kech = 13
        Case "01A": K2 = 3: Kech = 14
        Case "02A": K2 = 4: Kech = 15
        Case Else: K2 = 5: Kech = 16
    End Select
'_________________________________________________________________________________________________________
    If blnT2 Then
        For K = 1 To K2
            If blnPassif Then
                almT2(K).Mt1 = almT2(K).Mt1 + xDGAPPIS0.GAPPISMON
            Else
                almT2(K).Mt1 = almT2(K).Mt1 + xDGAPPIS0.GAPPISMON
            End If
        Next K
        
        If xDGAPPIS0.GAPPISTVF = "V" Then
            If blnPassif Then
                almT2(Kech).Mt4 = almT2(Kech).Mt4 + xDGAPPIS0.GAPPISMON
            Else
                almT2(Kech).Mt2 = almT2(Kech).Mt2 + xDGAPPIS0.GAPPISMON
            End If
        Else
            If blnPassif Then
                almT2(Kech).Mt3 = almT2(Kech).Mt3 + xDGAPPIS0.GAPPISMON
            Else
                almT2(Kech).Mt1 = almT2(Kech).Mt1 + xDGAPPIS0.GAPPISMON
            End If
        End If
        
        Kech = Kech + 10
        almT2(Kech).Mt1 = almT2(Kech).Mt1 + xDGAPPIS0.GAPPISTP1
        almT2(Kech).Mt2 = almT2(Kech).Mt2 + xDGAPPIS0.GAPPISTP2
        almT2(Kech).Mt3 = almT2(Kech).Mt3 + xDGAPPIS0.GAPPISTM1
        almT2(Kech).Mt4 = almT2(Kech).Mt4 + xDGAPPIS0.GAPPISTM2
        almT2(Kech).Mt5 = almT2(Kech).Mt5 + xDGAPPIS0.GAPPISMAR
        '=================
        GoTo DGAPPIS0_Next
        '=================
       
    End If
'==========================================================================================================
    blnT3 = False
    blnPassif = False
    
    If wGAPPISOPE = "CRE" And xDGAPPIS0.DGAPPISCLI = 11377 Then
        blnT3 = True: wT_Row = 31
        almT3(31).Mt1 = almT3(31).Mt1 + xDGAPPIS0.GAPPISMON
    End If
    If wGAPPISNAT = "DBQ" And xDGAPPIS0.DGAPPISCLI = 11012 Then
        blnPassif = True
        blnT3 = True: wT_Row = 41
        blnExclure_DBQ = True
        almT3(41).Mt1 = almT3(41).Mt1 + xDGAPPIS0.GAPPISMON
    End If
   
    Select Case xDGAPPIS0.DGAPPISVEC
        Case "01M": K2 = 0: Kech = 11
        Case "03M": K2 = 1: Kech = 12
        Case "06M": K2 = 2: Kech = 13
        Case "01A": K2 = 3: Kech = 14
        Case "02A": K2 = 4: Kech = 15
        Case Else: K2 = 5: Kech = 16
    End Select
'_________________________________________________________________________________________________________
    If blnT3 Then
        For K = 1 To K2
            If blnPassif Then
                almT3(K).Mt1 = almT3(K).Mt1 + xDGAPPIS0.GAPPISMON
            Else
                almT3(K).Mt1 = almT3(K).Mt1 + xDGAPPIS0.GAPPISMON
            End If
        Next K
        
        If xDGAPPIS0.GAPPISTVF = "V" Then
            If blnPassif Then
                almT3(Kech).Mt4 = almT3(Kech).Mt4 + xDGAPPIS0.GAPPISMON
            Else
                almT3(Kech).Mt2 = almT3(Kech).Mt2 + xDGAPPIS0.GAPPISMON
            End If
        Else
            If blnPassif Then
                almT3(Kech).Mt3 = almT3(Kech).Mt3 + xDGAPPIS0.GAPPISMON
            Else
                almT3(Kech).Mt1 = almT3(Kech).Mt1 + xDGAPPIS0.GAPPISMON
            End If
        End If
        
        Kech = Kech + 10
        almT3(Kech).Mt1 = almT3(Kech).Mt1 + xDGAPPIS0.GAPPISTP1
        almT3(Kech).Mt2 = almT3(Kech).Mt2 + xDGAPPIS0.GAPPISTP2
        almT3(Kech).Mt3 = almT3(Kech).Mt3 + xDGAPPIS0.GAPPISTM1
        almT3(Kech).Mt4 = almT3(Kech).Mt4 + xDGAPPIS0.GAPPISTM2
        almT3(Kech).Mt5 = almT3(Kech).Mt5 + xDGAPPIS0.GAPPISMAR
        '=================
        GoTo DGAPPIS0_Next
        '=================

    End If
'==========================================================================================================
    blnT4 = False
    blnPassif = False
    
'$JPL 2015-11-12
    Select Case wGAPPISTPR
        Case "TITOPC"
                    blnT4 = True: wT_Row = 37
                    almT4(37).Mt1 = almT4(37).Mt1 + xDGAPPIS0.GAPPISMON
    
        Case "RISPAY": blnPassif = True
                    blnT4 = True: wT_Row = 47
                    almT4(47).Mt1 = almT4(47).Mt1 + xDGAPPIS0.GAPPISMON
    End Select
    
    If Not blnT4 Then
    
        Select Case wGAPPISOPE
            Case "NOS", "NOB"
                        blnT4 = True
                        If xDGAPPIS0.GAPPISSEN = "A" Then
                            wT_Row = 31
                            almT4(31).Mt1 = almT4(31).Mt1 + xDGAPPIS0.GAPPISMON
                        Else
                            blnPassif = True: wT_Row = 41
                            almT4(41).Mt1 = almT4(41).Mt1 + xDGAPPIS0.GAPPISMON
                        End If
            Case "LOR", "LOB"
                        blnT4 = True
                        If xDGAPPIS0.GAPPISSEN = "A" Then
                            wT_Row = 32
                            almT4(32).Mt1 = almT4(32).Mt1 + xDGAPPIS0.GAPPISMON
                        Else
                            blnPassif = True: wT_Row = 42
                            almT4(42).Mt1 = almT4(42).Mt1 + xDGAPPIS0.GAPPISMON
                        End If
            Case "PRE"
                        If wGAPPISNAT <> "OAT" Then
                           blnT4 = True: wT_Row = 33
                           almT4(33).Mt1 = almT4(33).Mt1 + xDGAPPIS0.GAPPISMON
                        End If
            Case "CRE"
                        If wGAPPISNAT = "OAT" Or wGAPPISNAT = "PTI" Then
                        Else
                            blnT4 = True
                             If wGAPPISTPR = "BQCR030" Then
                                wT_Row = 34
                                almT4(34).Mt1 = almT4(34).Mt1 + xDGAPPIS0.GAPPISMON
                             Else
                                wT_Row = 36
                                almT4(36).Mt1 = almT4(36).Mt1 + xDGAPPIS0.GAPPISMON
                             End If
                        End If
            Case "BDF"
                        blnT4 = True: wT_Row = 38
                           almT4(38).Mt1 = almT4(38).Mt1 + xDGAPPIS0.GAPPISMON
                        
            Case "CAV", "DOR", "CBO"
                        blnT4 = True
                        If xDGAPPIS0.GAPPISSEN = "A" Then
                            wT_Row = 35
                            almT4(35).Mt1 = almT4(35).Mt1 + xDGAPPIS0.GAPPISMON
                        Else
                            blnPassif = True: wT_Row = 45
                            almT4(45).Mt1 = almT4(45).Mt1 + xDGAPPIS0.GAPPISMON
                        End If
            Case "EMP"
                        blnPassif = True
                        blnT4 = True: wT_Row = 43
                            almT4(43).Mt1 = almT4(43).Mt1 + xDGAPPIS0.GAPPISMON
            Case "EM1"
                        If Not blnExclure_DBQ Then
                            blnPassif = True
                            blnT4 = True
                            If wGAPPISNAT = "DBQ" Or wGAPPISNAT = "GEN" Then
                                wT_Row = 44
                                almT4(44).Mt1 = almT4(44).Mt1 + xDGAPPIS0.GAPPISMON
                            Else
                                wT_Row = 46
                               almT4(46).Mt1 = almT4(46).Mt1 + xDGAPPIS0.GAPPISMON
                            End If
                        End If
        End Select
   End If
   
   If Not blnT4 Then
        If Mid$(wGAPPISRUB, 1, 1) = "6" Or Mid$(wGAPPISRUB, 1, 1) = "7" Then
            blnPassif = True
            blnT4 = True: wT_Row = 48
            almT4(48).Mt1 = almT4(48).Mt1 + xDGAPPIS0.GAPPISMON
        End If
   
   End If
   
    Select Case xDGAPPIS0.DGAPPISVEC
        Case "01M": K2 = 0: Kech = 11
        Case "03M": K2 = 1: Kech = 12
        Case "06M": K2 = 2: Kech = 13
        Case "01A": K2 = 3: Kech = 14
        Case "02A": K2 = 4: Kech = 15
        Case Else: K2 = 5: Kech = 16
    End Select
'_________________________________________________________________________________________________________
    If blnT4 Then
        For K = 1 To K2
            If blnPassif Then
                almT4(K).Mt1 = almT4(K).Mt1 + xDGAPPIS0.GAPPISMON
            Else
                almT4(K).Mt1 = almT4(K).Mt1 + xDGAPPIS0.GAPPISMON
            End If
        Next K
        
        If xDGAPPIS0.GAPPISTVF = "V" Then
            If blnPassif Then
                almT4(Kech).Mt4 = almT4(Kech).Mt4 + xDGAPPIS0.GAPPISMON
            Else
                almT4(Kech).Mt2 = almT4(Kech).Mt2 + xDGAPPIS0.GAPPISMON
            End If
        Else
            If blnPassif Then
                almT4(Kech).Mt3 = almT4(Kech).Mt3 + xDGAPPIS0.GAPPISMON
            Else
                almT4(Kech).Mt1 = almT4(Kech).Mt1 + xDGAPPIS0.GAPPISMON
            End If
        End If
        
        Kech = Kech + 10
        almT4(Kech).Mt1 = almT4(Kech).Mt1 + xDGAPPIS0.GAPPISTP1
        almT4(Kech).Mt2 = almT4(Kech).Mt2 + xDGAPPIS0.GAPPISTP2
        almT4(Kech).Mt3 = almT4(Kech).Mt3 + xDGAPPIS0.GAPPISTM1
        almT4(Kech).Mt4 = almT4(Kech).Mt4 + xDGAPPIS0.GAPPISTM2
        almT4(Kech).Mt5 = almT4(Kech).Mt5 + xDGAPPIS0.GAPPISMAR
        '=================
        GoTo DGAPPIS0_Next
        '=================
        
    End If
'==========================================================================================================
DGAPPIS0_Next:
'=============
    mDetail_Row = mDetail_Row + 1
    
    If blnT1 Then
        wsExcel.Cells(mDetail_Row, 1) = "T1"
    Else
        If blnT2 Then
            wsExcel.Cells(mDetail_Row, 1) = "T2"
        Else
            If blnT3 Then
                wsExcel.Cells(mDetail_Row, 1) = "T3"
            Else
                If blnT4 Then wsExcel.Cells(mDetail_Row, 1) = "T4"
            End If
        End If
    End If
    Select Case wT_Row
        Case 0
        Case Is < 40: wsExcel.Cells(mDetail_Row, 2) = "A": wsExcel.Cells(mDetail_Row, 3) = wT_Row - 30
        Case Else: wsExcel.Cells(mDetail_Row, 2) = "P": wsExcel.Cells(mDetail_Row, 3) = wT_Row - 40
    End Select
    wsExcel.Cells(mDetail_Row, 4) = xDGAPPIS0.DGAPPISSEQ
    wsExcel.Cells(mDetail_Row, 5) = xDGAPPIS0.GAPPISMON
    wsExcel.Cells(mDetail_Row, 6) = xDGAPPIS0.GAPPISDEV
    wsExcel.Cells(mDetail_Row, 7) = mGAPPISMON
    wsExcel.Cells(mDetail_Row, 8) = xDGAPPIS0.GAPPISCLI
    wsExcel.Cells(mDetail_Row, 9) = xDGAPPIS0.GAPPISOPE
    wsExcel.Cells(mDetail_Row, 10) = xDGAPPIS0.GAPPISNAT
    wsExcel.Cells(mDetail_Row, 11) = xDGAPPIS0.GAPPISSEN
    wsExcel.Cells(mDetail_Row, 12) = xDGAPPIS0.GAPPISRUB
    wsExcel.Cells(mDetail_Row, 13) = xDGAPPIS0.GAPPISTPR
    wsExcel.Cells(mDetail_Row, 14) = xDGAPPIS0.GAPPISTVF

'==========================================================================================================
    
    rsSab.MoveNext
Loop

'__________________________________________________________________________________
Set rsSab = Nothing

Call cmdSelect_DWH_ALM_almT1
Call cmdSelect_DWH_ALM_almT2
Call cmdSelect_DWH_ALM_almT3
Call cmdSelect_DWH_ALM_almT4

wbExcel.SaveAs wFile

wbExcel.Close

'____________________________________________________________________________________
appExcel.Quit

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents
Exit Sub
Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0

    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

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
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), SAB_MNUAut)

blnSetfocus = True
Form_Init


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










Public Sub cmdSelect_DWH_ALM_Init_Page(lTxt As String)
Dim K As Integer

Call lstErr_AddItem(lstErr, cmdContext, "cmdSelect_DWH_ALM_init.... : "): DoEvents

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignRight
    .WrapText = False ' True
    .Font.Size = 8
    .Font.Name = "Calibri"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 85

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14" & lTxt & ", arrêté au " & dateImp10(wAmjMax) _
                                & vbCr


wsExcel.PageSetup.CenterHorizontally = True


wsExcel.PageSetup.PrintTitleRows = "$A1:$J1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"


wsExcel.Columns(1).ColumnWidth = 40: wsExcel.Cells(1, 1) = "Libellé": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(2).ColumnWidth = 9: wsExcel.Cells(1, 2) = "Actif": wsExcel.Columns(2).NumberFormat = "[Blue]### ### ### ###;[Red]-### ### ### ###"
wsExcel.Columns(3).ColumnWidth = 9: wsExcel.Cells(1, 3) = "Passif": wsExcel.Columns(3).NumberFormat = "[Blue]### ### ### ###;[Red]-### ### ### ###"
wsExcel.Columns(4).ColumnWidth = 9: wsExcel.Cells(1, 4) = "": wsExcel.Columns(4).NumberFormat = "[Blue]### ### ### ###;[Red]-### ### ### ###"

wsExcel.Columns(5).ColumnWidth = 30: wsExcel.Cells(1, 5) = "Libellé": wsExcel.Columns(5).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(6).ColumnWidth = 9: wsExcel.Cells(1, 6) = "A.Fixe /NR": wsExcel.Columns(6).NumberFormat = "[Blue]### ### ### ###;[Red]-### ### ### ###"
wsExcel.Columns(7).ColumnWidth = 9: wsExcel.Cells(1, 7) = "A. Variable": wsExcel.Columns(7).NumberFormat = "[Blue]### ### ### ###;[Red]-### ### ### ###"
wsExcel.Columns(8).ColumnWidth = 9: wsExcel.Cells(1, 8) = "P. Fixe / NR": wsExcel.Columns(8).NumberFormat = "[Blue]### ### ### ###;[Red]-### ### ### ###"
wsExcel.Columns(9).ColumnWidth = 9: wsExcel.Cells(1, 9) = "P. Variable": wsExcel.Columns(9).NumberFormat = "[Blue]### ### ### ###;[Red]-### ### ### ###"
wsExcel.Columns(10).ColumnWidth = 9: wsExcel.Cells(1, 10) = "Gap T Fixe (P-A)": wsExcel.Columns(10).NumberFormat = "[Blue]### ### ### ###;[Red]-### ### ### ###"


wsExcel.Columns(11).ColumnWidth = 30: wsExcel.Cells(1, 11) = "Libellé": wsExcel.Columns(11).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(12).ColumnWidth = 9: wsExcel.Cells(1, 12) = "Taux +1%": wsExcel.Columns(12).NumberFormat = "[Blue]### ### ### ###;[Red]-### ### ### ###"
wsExcel.Columns(13).ColumnWidth = 9: wsExcel.Cells(1, 13) = "Taux +2%": wsExcel.Columns(13).NumberFormat = "[Blue]### ### ### ###;[Red]-### ### ### ###"
wsExcel.Columns(14).ColumnWidth = 9: wsExcel.Cells(1, 14) = "Taux -1%": wsExcel.Columns(14).NumberFormat = "[Blue]### ### ### ###;[Red]-### ### ### ###"
wsExcel.Columns(15).ColumnWidth = 9: wsExcel.Cells(1, 15) = "Taux -2%": wsExcel.Columns(15).NumberFormat = "[Blue]### ### ### ###;[Red]-### ### ### ###"
wsExcel.Columns(16).ColumnWidth = 9: wsExcel.Cells(1, 16) = "Marge": wsExcel.Columns(16).NumberFormat = "[Blue]### ### ### ###;[Red]-### ### ### ###"


wsExcel.Cells(1, 1).Interior.Color = mColor_Y1

For K = 1 To 16
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next K

wsExcel.Cells(11, 1) = "Libellé"
wsExcel.Cells(11, 2) = "Actif"
wsExcel.Cells(11, 5) = "Libellé"
wsExcel.Cells(11, 6) = "Passif"

For K = 1 To 6
    wsExcel.Cells(11, K).Interior.Color = mColor_GB
    wsExcel.Cells(11, K).Font.Color = mColor_Z0
Next K


wsExcel.Cells(20, 2).FormulaLocal = "=SOMME(B12:B19)": wsExcel.Cells(20, 2).Interior.Color = mColor_G1
wsExcel.Cells(20, 6).FormulaLocal = "=SOMME(F12:F19)": wsExcel.Cells(20, 6).Interior.Color = mColor_G1

'__________________________________________________________________________________



'======================================================================================================

Exit_sub:

End Sub

Public Sub cmdSelect_DWH_ALM_Init_Page_Detail(lTxt As String)
Dim K As Integer

Call lstErr_AddItem(lstErr, cmdContext, "cmdSelect_DWH_ALM_init.... : "): DoEvents

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignRight
    .WrapText = False ' True
    .Font.Size = 8
    .Font.Name = "Calibri"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 85

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14" & lTxt & ", arrêté au " & dateImp10(wAmjMax) _
                                & vbCr


wsExcel.PageSetup.CenterHorizontally = True


wsExcel.PageSetup.PrintTitleRows = "$A1:$J1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"


wsExcel.Columns(1).ColumnWidth = 7: wsExcel.Cells(1, 1) = "Tableau": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(2).ColumnWidth = 5: wsExcel.Cells(1, 2) = "A/P": wsExcel.Columns(2).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(3).ColumnWidth = 5: wsExcel.Cells(1, 3) = "Ligne": wsExcel.Columns(3).NumberFormat = "[Blue]###;[Red]-###"
wsExcel.Columns(4).ColumnWidth = 7: wsExcel.Cells(1, 4) = "GAP_Séq": wsExcel.Columns(4).NumberFormat = "[Blue]### ### ###;[Red]-### ### ###"
wsExcel.Columns(5).ColumnWidth = 15: wsExcel.Cells(1, 5) = "EUR": wsExcel.Columns(5).NumberFormat = "[Blue]### ### ### ###.##;[Red]-### ### ### ###.##"
wsExcel.Columns(6).ColumnWidth = 5: wsExcel.Cells(1, 6) = "Dev": wsExcel.Columns(6).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(7).ColumnWidth = 15: wsExcel.Cells(1, 7) = "Mt ": wsExcel.Columns(7).NumberFormat = "[Blue]### ### ### ###.##;[Red]-### ### ### ###.##"
wsExcel.Columns(8).ColumnWidth = 5: wsExcel.Cells(1, 8) = "CLI": wsExcel.Columns(8).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(9).ColumnWidth = 5: wsExcel.Cells(1, 9) = "OPE": wsExcel.Columns(9).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(10).ColumnWidth = 8: wsExcel.Cells(1, 10) = "NAT": wsExcel.Columns(10).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(11).ColumnWidth = 5: wsExcel.Cells(1, 11) = "SEN": wsExcel.Columns(11).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(12).ColumnWidth = 8: wsExcel.Cells(1, 12) = "RUB": wsExcel.Columns(12).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(13).ColumnWidth = 8: wsExcel.Cells(1, 13) = "TPR": wsExcel.Columns(13).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(14).ColumnWidth = 5: wsExcel.Cells(1, 14) = "TVF": wsExcel.Columns(14).HorizontalAlignment = Excel.xlHAlignLeft


'wsExcel.Cells(1, 1).Interior.Color = mColor_Y1

For K = 1 To 16
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next K
'__________________________________________________________________________________



'======================================================================================================

Exit_sub:

End Sub


Public Sub cmdSelect_DWH_ALM_Init()
Dim K As Integer

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "DWH_ALM"
    .Subject = "DWH_ALM"
End With

'__________________________________________________________________________________

appExcel.Worksheets.Add
appExcel.Worksheets.Add
appExcel.Worksheets.Add
appExcel.Worksheets.Add
appExcel.Worksheets.Add

Set wsExcel = wbExcel.Sheets(1): wsExcel.Name = "Global"
Set wsExcel = wbExcel.Sheets(2): wsExcel.Name = "T1"
Set wsExcel = wbExcel.Sheets(3): wsExcel.Name = "T2"
Set wsExcel = wbExcel.Sheets(4): wsExcel.Name = "T3"
Set wsExcel = wbExcel.Sheets(5): wsExcel.Name = "T4"
Set wsExcel = wbExcel.Sheets(6): wsExcel.Name = "Détail"

Set wsExcel = wbExcel.Sheets(1)
cmdSelect_DWH_ALM_Init_Page ("Tableau5.Périmètre global")

Set wsExcel = wbExcel.Sheets(2)
cmdSelect_DWH_ALM_Init_Page ("Tableau1.Fonds propres et actifs alloués")

Set wsExcel = wbExcel.Sheets(3)
cmdSelect_DWH_ALM_Init_Page ("Tableau2.Micro-adossement : gage espèces / banque de France + titres")

Set wsExcel = wbExcel.Sheets(4)
cmdSelect_DWH_ALM_Init_Page ("Tableau3.Micro-adossement : opération financière cER")

Set wsExcel = wbExcel.Sheets(5)
cmdSelect_DWH_ALM_Init_Page ("Tableau4.autres positions au bilan non micro-adossées et excédent banque de France")

Set wsExcel = wbExcel.Sheets(6)
cmdSelect_DWH_ALM_Init_Page_Detail ("Détail des affectations")

'==========================================================================================
For K = 1 To 50
    almT1(K).Mt1 = 0: almT1(K).Mt2 = 0: almT1(K).Mt3 = 0: almT1(K).Mt4 = 0: almT1(K).Mt5 = 0
Next K

almT1(1).Lib = "Risque de liquidité : < 1 mois": almT1(1).Row = 2: almT1(1).Col = 1
almT1(2).Lib = "Risque de liquidité : 1-3 mois": almT1(2).Row = 3: almT1(2).Col = 1
almT1(3).Lib = "Risque de liquidité : 3-6 mois": almT1(3).Row = 4: almT1(3).Col = 1
almT1(4).Lib = "Risque de liquidité : 6-12 mois": almT1(4).Row = 5: almT1(4).Col = 1
almT1(5).Lib = "Risque de liquidité : 12-24 mois": almT1(5).Row = 6: almT1(5).Col = 1
almT1(6).Lib = "Risque de liquidité : > 24 mois": almT1(6).Row = 7: almT1(6).Col = 1

almT1(11).Lib = "Risque de taux 1 : < 1 mois": almT1(11).Row = 2: almT1(11).Col = 5
almT1(12).Lib = "Risque de taux 1 : 1-3 mois": almT1(12).Row = 3: almT1(12).Col = 5
almT1(13).Lib = "Risque de taux 1 : 3-6 mois": almT1(13).Row = 4: almT1(13).Col = 5
almT1(14).Lib = "Risque de taux 1 : 6-12 mois": almT1(14).Row = 5: almT1(14).Col = 5
almT1(15).Lib = "Risque de taux 1 : 12-24 mois": almT1(15).Row = 6: almT1(15).Col = 5
almT1(16).Lib = "Risque de taux 1 : > 24 mois": almT1(16).Row = 7: almT1(16).Col = 5

almT1(21).Lib = "Risque de taux 2 : < 1 mois": almT1(21).Row = 2: almT1(21).Col = 11
almT1(22).Lib = "Risque de taux 2 : 1-3 mois": almT1(22).Row = 3: almT1(22).Col = 11
almT1(23).Lib = "Risque de taux 2 : 3-6 mois": almT1(23).Row = 4: almT1(23).Col = 11
almT1(24).Lib = "Risque de taux 2 : 6-12 mois": almT1(24).Row = 5: almT1(24).Col = 11
almT1(25).Lib = "Risque de taux 2 : 12-24 mois": almT1(25).Row = 6: almT1(25).Col = 11
almT1(26).Lib = "Risque de taux 2 : > 24 mois": almT1(26).Row = 7: almT1(26).Col = 11

almT1(31).Lib = "Immobilisations non amorties": almT1(31).Row = 12: almT1(31).Col = 1
almT1(32).Lib = "Créances douteuses nettes de provisions": almT1(32).Row = 13: almT1(32).Col = 1
almT1(33).Lib = "Titres de participation nets de provisions": almT1(33).Row = 14: almT1(33).Col = 1
almT1(34).Lib = "Titres de placement Actions nets de provisions": almT1(34).Row = 15: almT1(34).Col = 1
almT1(35).Lib = "Titres de placement obligations nets de provisions": almT1(35).Row = 16: almT1(35).Col = 1
almT1(36).Lib = "Litiges fiscaux nets de provisions": almT1(36).Row = 17: almT1(36).Col = 1

almT1(41).Lib = "Fonds propres": almT1(41).Row = 12: almT1(41).Col = 5
'''almT1(42).Lib = "Provisions risques-pays": almT1(42).Row = 13: almT1(42).Col = 5

'==========================================================================================
For K = 1 To 50
    almT2(K).Mt1 = 0: almT2(K).Mt2 = 0: almT2(K).Mt3 = 0: almT2(K).Mt4 = 0: almT2(K).Mt5 = 0
Next K
For K = 1 To 30
    almT2(K).Lib = almT1(K).Lib: almT2(K).Row = almT1(K).Row: almT2(K).Col = almT1(K).Col
Next K

almT2(31).Lib = "Banque de France": almT2(31).Row = 12: almT2(31).Col = 1
almT2(32).Lib = "FCP Palatine": almT2(32).Row = 13: almT2(32).Col = 1
almT2(33).Lib = "Titres de trésorerie OCT (obligataires court terme)": almT2(33).Row = 14: almT2(33).Col = 1
almT2(34).Lib = "Autres obligations": almT2(34).Row = 15: almT2(34).Col = 1

almT2(41).Lib = "Gage-espèces BEA résiduel": almT2(41).Row = 12: almT2(41).Col = 5
almT2(42).Lib = "Gage-espèces LFB résiduel": almT2(42).Row = 13: almT2(42).Col = 5
'==========================================================================================
For K = 1 To 50
    almT3(K).Mt1 = 0: almT3(K).Mt2 = 0: almT3(K).Mt3 = 0: almT3(K).Mt4 = 0: almT3(K).Mt5 = 0
Next K
For K = 1 To 30
    almT3(K).Lib = almT1(K).Lib: almT3(K).Row = almT1(K).Row: almT3(K).Col = almT1(K).Col
Next K

almT3(31).Lib = "Crédit d'équipement (Financière CER)": almT3(31).Row = 12: almT3(31).Col = 1

almT3(41).Lib = "Dépôt de garantie LFB": almT3(41).Row = 12: almT3(41).Col = 5
'==========================================================================================
For K = 1 To 50
    almT4(K).Mt1 = 0: almT4(K).Mt2 = 0: almT4(K).Mt3 = 0: almT4(K).Mt4 = 0: almT4(K).Mt5 = 0
Next K
For K = 1 To 30
    almT4(K).Lib = almT1(K).Lib: almT4(K).Row = almT1(K).Row: almT4(K).Col = almT1(K).Col
Next K

almT4(31).Lib = "Nostri": almT4(31).Row = 12: almT4(31).Col = 1
almT4(32).Lib = "Lori": almT4(32).Row = 13: almT4(32).Col = 1
almT4(33).Lib = "Prêts interbancaires": almT4(33).Row = 14: almT4(33).Col = 1
almT4(34).Lib = "Créances bancaires": almT4(34).Row = 15: almT4(34).Col = 1
almT4(35).Lib = "Comptes à vue débiteurs": almT4(35).Row = 16: almT4(35).Col = 1
almT4(36).Lib = "Créances commerciales": almT4(36).Row = 17: almT4(36).Col = 1
almT4(37).Lib = "TITOPC": almT4(37).Row = 18: almT4(37).Col = 1
almT4(38).Lib = "Solde Banque de France": almT4(38).Row = 19: almT4(38).Col = 1

almT4(41).Lib = "Nostri": almT4(41).Row = 12: almT4(41).Col = 5
almT4(42).Lib = "Lori": almT4(42).Row = 13: almT4(42).Col = 5
almT4(43).Lib = "Emprunts  interbancaires": almT4(43).Row = 14: almT4(43).Col = 5
almT4(44).Lib = "Dépôts de garantie": almT4(44).Row = 15: almT4(44).Col = 5
almT4(45).Lib = "Comptes à vue créditeurs": almT4(45).Row = 16: almT4(45).Col = 5
almT4(46).Lib = "Dépôts à terme": almT4(46).Row = 17: almT4(46).Col = 5
almT4(47).Lib = "Provisions risques-pays": almT4(47).Row = 18: almT4(47).Col = 5
almT4(48).Lib = "Autres et résultat": almT4(48).Row = 19: almT4(48).Col = 5



End Sub

Public Sub cmdSelect_DWH_ALM_almT1()
Dim K As Integer
Call lstErr_AddItem(lstErr, cmdContext, "cmdSelect_DWH_ALM_almT1.... : "): DoEvents

Set wsExcel = wbExcel.Sheets(2)

For K = 1 To 50
    If almT1(K).Row <> 0 Then
        wsExcel.Cells(almT1(K).Row, almT1(K).Col) = almT1(K).Lib
        If almT1(K).Mt1 <> 0 Then wsExcel.Cells(almT1(K).Row, almT1(K).Col + 1) = almT1(K).Mt1 / 1000
        If almT1(K).Mt2 <> 0 Then wsExcel.Cells(almT1(K).Row, almT1(K).Col + 2) = almT1(K).Mt2 / 1000
        If almT1(K).Mt3 <> 0 Then wsExcel.Cells(almT1(K).Row, almT1(K).Col + 3) = almT1(K).Mt3 / 1000
        If almT1(K).Mt4 <> 0 Then wsExcel.Cells(almT1(K).Row, almT1(K).Col + 4) = almT1(K).Mt4 / 1000
        If almT1(K).Mt5 <> 0 Then wsExcel.Cells(almT1(K).Row, almT1(K).Col + 5) = almT1(K).Mt5 / 1000
   
    End If
Next K
End Sub
Public Sub cmdSelect_DWH_ALM_almT2()
Dim K As Integer
Call lstErr_AddItem(lstErr, cmdContext, "cmdSelect_DWH_ALM_almT2.... : "): DoEvents

Set wsExcel = wbExcel.Sheets(3)

For K = 1 To 50
    If almT2(K).Row <> 0 Then
        wsExcel.Cells(almT2(K).Row, almT2(K).Col) = almT2(K).Lib
        If almT2(K).Mt1 <> 0 Then wsExcel.Cells(almT2(K).Row, almT2(K).Col + 1) = almT2(K).Mt1 / 1000
        If almT2(K).Mt2 <> 0 Then wsExcel.Cells(almT2(K).Row, almT2(K).Col + 2) = almT2(K).Mt2 / 1000
        If almT2(K).Mt3 <> 0 Then wsExcel.Cells(almT2(K).Row, almT2(K).Col + 3) = almT2(K).Mt3 / 1000
        If almT2(K).Mt4 <> 0 Then wsExcel.Cells(almT2(K).Row, almT2(K).Col + 4) = almT2(K).Mt4 / 1000
        If almT2(K).Mt5 <> 0 Then wsExcel.Cells(almT2(K).Row, almT2(K).Col + 5) = almT2(K).Mt5 / 1000
   
    End If
Next K

wsExcel.Cells(12, 2).FormulaLocal = "= - SOMME(B13:B19) - SOMME(F12:F19)": wsExcel.Cells(12, 2).Interior.Color = mColor_Y1

End Sub

Public Sub cmdSelect_DWH_ALM_almT3()
Dim K As Integer
Call lstErr_AddItem(lstErr, cmdContext, "cmdSelect_DWH_ALM_almT3.... : "): DoEvents

Set wsExcel = wbExcel.Sheets(4)

For K = 1 To 50
    If almT3(K).Row <> 0 Then
        wsExcel.Cells(almT3(K).Row, almT3(K).Col) = almT3(K).Lib
        If almT3(K).Mt1 <> 0 Then wsExcel.Cells(almT3(K).Row, almT3(K).Col + 1) = almT3(K).Mt1 / 1000
        If almT3(K).Mt2 <> 0 Then wsExcel.Cells(almT3(K).Row, almT3(K).Col + 2) = almT3(K).Mt2 / 1000
        If almT3(K).Mt3 <> 0 Then wsExcel.Cells(almT3(K).Row, almT3(K).Col + 3) = almT3(K).Mt3 / 1000
        If almT3(K).Mt4 <> 0 Then wsExcel.Cells(almT3(K).Row, almT3(K).Col + 4) = almT3(K).Mt4 / 1000
        If almT3(K).Mt5 <> 0 Then wsExcel.Cells(almT3(K).Row, almT3(K).Col + 5) = almT3(K).Mt5 / 1000
   
    End If
Next K
End Sub

Public Sub cmdSelect_DWH_ALM_almT4()
Dim K As Integer
Call lstErr_AddItem(lstErr, cmdContext, "cmdSelect_DWH_ALM_almT4.... : "): DoEvents

Set wsExcel = wbExcel.Sheets(5)

For K = 1 To 50
    If almT4(K).Row <> 0 Then
        wsExcel.Cells(almT4(K).Row, almT4(K).Col) = almT4(K).Lib
        If almT4(K).Mt1 <> 0 Then wsExcel.Cells(almT4(K).Row, almT4(K).Col + 1) = almT4(K).Mt1 / 1000
        If almT4(K).Mt2 <> 0 Then wsExcel.Cells(almT4(K).Row, almT4(K).Col + 2) = almT4(K).Mt2 / 1000
        If almT4(K).Mt3 <> 0 Then wsExcel.Cells(almT4(K).Row, almT4(K).Col + 3) = almT4(K).Mt3 / 1000
        If almT4(K).Mt4 <> 0 Then wsExcel.Cells(almT4(K).Row, almT4(K).Col + 4) = almT4(K).Mt4 / 1000
        If almT4(K).Mt5 <> 0 Then wsExcel.Cells(almT4(K).Row, almT4(K).Col + 5) = almT4(K).Mt5 / 1000
   
    End If
Next K

wsExcel.Cells(10, 1) = "GAPPISRUB = 'BDF' (cellule de calcul: ne pas effacer)"
wsExcel.Cells(10, 2) = almT4(38).Mt1 / 1000
wsExcel.Cells(10, 2).Interior.Color = RGB(196, 196, 196)
wsExcel.Cells(19, 2).FormulaLocal = "= 'T4'!B10 -  'T2'!B12": wsExcel.Cells(19, 2).Interior.Color = mColor_Y1

End Sub


