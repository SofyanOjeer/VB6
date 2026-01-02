VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDCRETRO 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA_DWH"
   ClientHeight    =   9144
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   13560
   Icon            =   "DCRETRO.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9144
   ScaleWidth      =   13560
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   432
      Left            =   7800
      TabIndex        =   15
      Top             =   0
      Width           =   5175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   14
      Top             =   480
      Width           =   13530
      _ExtentX        =   23855
      _ExtentY        =   15261
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Gestion sur les mois traités"
      TabPicture(0)   =   "DCRETRO.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "....."
      TabPicture(1)   =   "DCRETRO.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   13290
         Begin VB.CommandButton cmdSelect_Ajouter 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Ajouter"
            Enabled         =   0   'False
            Height          =   285
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   720
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Frame fraDCRETRO 
            BackColor       =   &H80000013&
            Caption         =   "Gestion du fichier - DCRETRO -"
            Height          =   5775
            Left            =   6480
            TabIndex        =   27
            Top             =   1800
            Visible         =   0   'False
            Width           =   6015
            Begin VB.TextBox txtfg_DRETCLR 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2160
               TabIndex        =   43
               Text            =   "Client renta"
               Top             =   3600
               Width           =   855
            End
            Begin VB.TextBox txtfg_DRETREF 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2760
               TabIndex        =   37
               Text            =   "Référence"
               Top             =   2280
               Width           =   1935
            End
            Begin VB.TextBox txtfg_DRETEVT 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2160
               TabIndex        =   36
               Text            =   "Event"
               Top             =   2280
               Width           =   495
            End
            Begin VB.TextBox txtfg_DRETNAT 
               Enabled         =   0   'False
               Height          =   300
               Left            =   4800
               TabIndex        =   39
               Text            =   "Nature"
               Top             =   2280
               Width           =   855
            End
            Begin VB.TextBox txtfg_DRETSEQ 
               Height          =   300
               Left            =   3600
               TabIndex        =   35
               Text            =   "Seq"
               Top             =   1680
               Width           =   375
            End
            Begin VB.TextBox txtfg_DRETNUM 
               Height          =   300
               Left            =   3000
               TabIndex        =   29
               Text            =   "No opération"
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox txtfg_DRETOPE 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2160
               TabIndex        =   28
               Text            =   "Opération"
               Top             =   1320
               Width           =   495
            End
            Begin VB.TextBox txtfg_DRETSER 
               Enabled         =   0   'False
               Height          =   300
               Left            =   3120
               TabIndex        =   23
               Text            =   "Service"
               Top             =   960
               Width           =   375
            End
            Begin VB.TextBox txtfg_DRETAGE 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2640
               TabIndex        =   22
               Text            =   "Agence"
               Top             =   960
               Width           =   375
            End
            Begin VB.TextBox txtfg_DRETSSE 
               Enabled         =   0   'False
               Height          =   300
               Left            =   3600
               TabIndex        =   24
               Text            =   "Sous-service"
               Top             =   960
               Width           =   375
            End
            Begin VB.ComboBox cbo_CDEV 
               Height          =   315
               Left            =   3360
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   2640
               Width           =   2295
            End
            Begin VB.TextBox txtfg_DRETMNT1 
               Height          =   300
               Left            =   2160
               TabIndex        =   41
               Text            =   "Montant comm 1"
               Top             =   3000
               Width           =   1695
            End
            Begin VB.ComboBox cbo_CRTA 
               Height          =   315
               Left            =   2160
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   44
               Top             =   3960
               Width           =   2775
            End
            Begin VB.CommandButton cmdfra_Creer 
               BackColor       =   &H00FFC0FF&
               Caption         =   "Créer"
               Height          =   495
               Left            =   5040
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   5040
               Width           =   855
            End
            Begin VB.CommandButton cmdfra_Modifier 
               BackColor       =   &H00FF80FF&
               Caption         =   "Modifier"
               Height          =   495
               Left            =   3960
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   5040
               Width           =   855
            End
            Begin VB.CommandButton cmdfra_Supprimer 
               BackColor       =   &H000000FF&
               Caption         =   "Supprimer"
               Height          =   495
               Left            =   2760
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   5040
               Width           =   855
            End
            Begin VB.TextBox txtfg_DRETMNT2 
               Height          =   300
               Left            =   3960
               TabIndex        =   42
               Text            =   "Montant comm 2"
               Top             =   3000
               Width           =   1695
            End
            Begin VB.TextBox txtfg_DRETCTG 
               Height          =   300
               Left            =   2160
               TabIndex        =   45
               Text            =   "Comptage"
               Top             =   4320
               Width           =   615
            End
            Begin VB.TextBox txtfg_DRETCLI 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2160
               TabIndex        =   38
               Text            =   "No client"
               Top             =   2640
               Width           =   855
            End
            Begin VB.TextBox txtfg_DRETETB 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2160
               TabIndex        =   21
               Text            =   "Etablissement"
               Top             =   960
               Width           =   375
            End
            Begin VB.TextBox txtfg_DRETVER 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2640
               TabIndex        =   12
               Text            =   "Version"
               Top             =   600
               Width           =   375
            End
            Begin VB.TextBox txtfg_DRETSTA 
               Height          =   300
               Left            =   2160
               TabIndex        =   11
               Text            =   "Statut"
               Top             =   600
               Width           =   375
            End
            Begin MSComCtl2.DTPicker txtfg_DRETDTR 
               Height          =   300
               Left            =   2160
               TabIndex        =   34
               Top             =   1680
               Width           =   1335
               _ExtentX        =   2350
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   10223619
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblfg_DRETCLR 
               BackColor       =   &H80000013&
               Caption         =   "Client renta"
               Height          =   300
               Left            =   240
               TabIndex        =   8
               Top             =   3600
               Width           =   1815
            End
            Begin VB.Label lblfg_DRETEVT_REF 
               BackColor       =   &H80000013&
               Caption         =   "Event / Réf / Nature"
               Height          =   300
               Left            =   240
               TabIndex        =   5
               Top             =   2280
               Width           =   1695
            End
            Begin VB.Label lblfg_DRETDTR_SEQ 
               BackColor       =   &H80000013&
               Caption         =   "Date traitement / Seq"
               Height          =   300
               Left            =   240
               TabIndex        =   4
               Top             =   1680
               Width           =   1815
            End
            Begin VB.Label lblfg_DRETOPE_NUM 
               BackColor       =   &H80000013&
               Caption         =   "Opération / no opération"
               Height          =   300
               Left            =   240
               TabIndex        =   3
               Top             =   1320
               Width           =   1815
            End
            Begin VB.Label lblfg_DRETMNT 
               BackColor       =   &H80000013&
               Caption         =   "Montants commissions "
               Height          =   375
               Left            =   240
               TabIndex        =   7
               Top             =   3000
               Width           =   1815
            End
            Begin VB.Label lblfg_DRTACTR 
               BackColor       =   &H80000013&
               Caption         =   "Comptage"
               Height          =   300
               Left            =   240
               TabIndex        =   10
               Top             =   4320
               Width           =   1695
            End
            Begin VB.Label lblfg_DRTACRTA 
               BackColor       =   &H80000013&
               Caption         =   "Code renta"
               Height          =   300
               Left            =   240
               TabIndex        =   9
               Top             =   3960
               Width           =   1815
            End
            Begin VB.Label lblfg_DRETCLI_DEV 
               BackColor       =   &H80000013&
               Caption         =   "Client / Devise"
               Height          =   300
               Left            =   240
               TabIndex        =   6
               Top             =   2640
               Width           =   1815
            End
            Begin VB.Label lblfg_DRETETB 
               BackColor       =   &H80000013&
               Caption         =   "Etab - Agce - SER - SSE"
               Height          =   300
               Left            =   240
               TabIndex        =   2
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label lblfg_DRETSTA_VER 
               BackColor       =   &H80000013&
               Caption         =   "Statut / Version"
               Height          =   300
               Left            =   240
               TabIndex        =   1
               Top             =   600
               Width           =   1815
            End
         End
         Begin VB.Frame fraSelect_Options 
            Height          =   1005
            Left            =   240
            TabIndex        =   19
            Top             =   120
            Width           =   11355
            Begin MSComCtl2.DTPicker txtSelect_DRETPER 
               Height          =   300
               Left            =   1920
               TabIndex        =   20
               Top             =   240
               Width           =   1455
               _ExtentX        =   2561
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   10223619
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblSelect_DRTAPER 
               Caption         =   "Période de traitement"
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Enabled         =   0   'False
            Height          =   315
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   240
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6825
            Left            =   0
            TabIndex        =   25
            Top             =   1200
            Width           =   12840
            _ExtentX        =   22648
            _ExtentY        =   12044
            _Version        =   393216
            Rows            =   1
            Cols            =   23
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
            FormatString    =   $"DCRETRO.frx":047A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.4
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
      TabIndex        =   13
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13080
      Picture         =   "DCRETRO.frx":0560
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
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   16
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
End
Attribute VB_Name = "frmDCRETRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit

Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim BIA_DWH_Aut As typeAuthorization
Dim curX1 As Currency, curX2 As Currency
Dim blnAuto As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim cnADO As New ADODB.Connection, rsado As New ADODB.Recordset, errADO As ADODB.Error
Dim blnTransaction As Boolean

'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wAmjDTR As String, wHmsMin As Long, wHmsMax As Long
Dim xDCRETRO As typeDCRETRO, newDCRETRO As typeDCRETRO, oldDCRETRO As typeDCRETRO
Dim arrDCRETRO() As typeDCRETRO, arrDCRETRO_Nb As Long, arrDCRETRO_Max As Long, arrDCRETRO_Index As Long
Dim xDRTAGRP As typeDRTAGRP
Dim xWhere_CRTA As String
Dim xDDEVISE As typeDDEVISE
Dim xWhere_CDEV As String

'______________________________________________________________________

Dim fraDCRETRO_FormatString As String, fraDCRETRO_K As Integer
Dim fraDCRETRO_RowDisplay As Integer, fraDCRETRO_RowClick As Integer, fraDCRETRO_ColClick As Integer
Dim fraDCRETRO_ColorClick As Long, fraDCRETRO_ColorDisplay As Long
Dim fraDCRETRO_Sort1 As Integer, fraDCRETRO_Sort2 As Integer
Dim fraDCRETRO_SortAD As Integer, fraDCRETRO_Sort1_Old As Integer
Dim fraDCRETRO_arrIndex As Integer
Dim blnfraDCRETRO_DisplayLine As Boolean

Dim meDCRETRO_Status As typeDCRETRO, oldDCRETRO_Status As typeDCRETRO

Dim rsADO_Update As New ADODB.Recordset
Dim mTitleText As String

Dim fgSelect_TopRow_Memo As Integer









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
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgselect_Display"
    
For I = 1 To arrDCRETRO_Nb
         
    xDCRETRO = arrDCRETRO(I)
    If xDCRETRO.DRETSTA <> "" Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
    End If
Next I

fgSelect.Visible = True
fraDCRETRO.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True

Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrDCRETRO_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
    
End Sub

Private Sub arrDCRETRO_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrDCRETRO(101)
arrDCRETRO_Max = 100: arrDCRETRO_Nb = 0

Set rsado = Nothing

xSql = "select * from " & paramIBM_Library_BODWH & ".DCRETRO " & xWhere & " order by DRETVER, DRETETB, DRETAGE, DRETSER, DRETSSE, DRETOPE, DRETNUM, DRETDTR, DRETSEQ"
Set rsado = cnADO.Execute(xSql)

Do While Not rsado.EOF
    V = srvDCRETRO_GetBuffer_ODBC(rsado, xDCRETRO)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDCRETRO.fgselect_Display"
        '' Exit Sub
     Else
         arrDCRETRO_Nb = arrDCRETRO_Nb + 1
         If arrDCRETRO_Nb > arrDCRETRO_Max Then
             arrDCRETRO_Max = arrDCRETRO_Max + 50
             ReDim Preserve arrDCRETRO(arrDCRETRO_Max)
         End If
         
         arrDCRETRO(arrDCRETRO_Nb) = xDCRETRO
    End If
    rsado.MoveNext
Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub DRTAGRP_CRTA_SQL(xWhere_CRTA As String)

Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler

cbo_CRTA.Clear

Set rsado = Nothing

xSql = "select * from " & paramIBM_Library_BODWH & ".DRTAGRP " & xWhere_CRTA & " order by DRGRCRTA"
Set rsado = cnADO.Execute(xSql)

Do While Not rsado.EOF
    V = srvDRTAGRP_GetBuffer_ODBC(rsado, xDRTAGRP)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDCRETRO.DRTAGRP_CRTA_SQL"
        '' Exit Sub
     Else
         cbo_CRTA.AddItem xDRTAGRP.DRGRCRTA & "  " & xDRTAGRP.DRGRLIB
    End If
    rsado.MoveNext
Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub





Private Sub DDEVISE_CDEV_SQL(xWhere_CDEV As String)

Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler

cbo_CDEV.Clear

Set rsado = Nothing

xSql = "select * from " & paramIBM_Library_BODWH & ".DDEVISE " & xWhere_CDEV & " order by DDEVDEV, DDEVLIB"
Set rsado = cnADO.Execute(xSql)

Do While Not rsado.EOF
    V = srvDDEVISE_GetBuffer_ODBC(rsado, xDDEVISE)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDCRETRO.DDEVISE_CDEV_SQL"
        '' Exit Sub
     Else
         cbo_CDEV.AddItem xDDEVISE.DDEVDEV & "  " & xDDEVISE.DDEVLIB
    End If
    rsado.MoveNext
Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub





Public Sub fgSelect_DisplayLine(lIndex As Long)

On Error Resume Next

fgSelect.Col = 0: fgSelect.Text = xDCRETRO.DRETSTA
fgSelect.Col = 1: fgSelect.Text = xDCRETRO.DRETVER
fgSelect.Col = 2: fgSelect.Text = dateImp10(xDCRETRO.DRETPER)
fgSelect.Col = 3: fgSelect.Text = xDCRETRO.DRETETB
fgSelect.Col = 4: fgSelect.Text = xDCRETRO.DRETAGE
fgSelect.Col = 5: fgSelect.Text = xDCRETRO.DRETSER
fgSelect.Col = 6: fgSelect.Text = xDCRETRO.DRETSSE
fgSelect.Col = 7: fgSelect.Text = xDCRETRO.DRETOPE
fgSelect.Col = 8: fgSelect.Text = xDCRETRO.DRETNUM
fgSelect.Col = 9: fgSelect.Text = dateImp10(xDCRETRO.DRETDTR)
fgSelect.Col = 10: fgSelect.Text = xDCRETRO.DRETSEQ
fgSelect.Col = 11: fgSelect.Text = xDCRETRO.DRETEVT
fgSelect.Col = 12: fgSelect.Text = xDCRETRO.DRETNAT
fgSelect.Col = 13: fgSelect.Text = xDCRETRO.DRETREF
fgSelect.Col = 14: fgSelect.Text = xDCRETRO.DRETCLI
fgSelect.Col = 15: fgSelect.Text = xDCRETRO.DRETDEV

fgSelect.Col = 16
If xDCRETRO.DRETMNT1 = 0 Then
    fgSelect.Text = ""
Else
    fgSelect.Text = Format$(xDCRETRO.DRETMNT1, "### ### ### ##0.00")
    ' Mettre en rouge les montants débiteurs
    'If xDCRETRO.DRETMNT1 >= 0 Then
    '    fgSelect.CellForeColor = vbBlue
    'Else
    '    fgSelect.CellForeColor = vbRed
    'End If
End If
fgSelect.Col = 17
If xDCRETRO.DRETMNT2 = 0 Then
    fgSelect.Text = ""
Else
    fgSelect.Text = Format$(xDCRETRO.DRETMNT2, "### ### ### ##0.00")
    ' Mettre en rouge les montants débiteurs
    'If xDCRETRO.DRETMNT2 >= 0 Then
    '    fgSelect.CellForeColor = vbBlue
    'Else
    '    fgSelect.CellForeColor = vbRed
    'End If
End If

fgSelect.Col = 18: fgSelect.Text = xDCRETRO.DRETCLR
fgSelect.Col = 19: fgSelect.Text = xDCRETRO.DRETCRTA
fgSelect.Col = 20: fgSelect.Text = xDCRETRO.DRETCTG

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

Public Sub cmdDCRETRO_Charger()

srvDCRETRO_Init newDCRETRO

newDCRETRO.DRETSTA = Trim(txtfg_DRETSTA)
newDCRETRO.DRETVER = Val(txtfg_DRETVER)
newDCRETRO.DRETPER = wAmjMax

newDCRETRO.DRETETB = Trim(txtfg_DRETETB)
newDCRETRO.DRETAGE = Trim(txtfg_DRETAGE)
newDCRETRO.DRETSER = Trim(txtfg_DRETSER)
newDCRETRO.DRETSSE = Trim(txtfg_DRETSSE)

newDCRETRO.DRETOPE = Trim(txtfg_DRETOPE)
newDCRETRO.DRETNUM = Val(Mid$(txtfg_DRETNUM, 1, 9))
Call DTPicker_Control(txtfg_DRETDTR, wAmjDTR)
newDCRETRO.DRETDTR = wAmjDTR
newDCRETRO.DRETSEQ = Val(Mid$(txtfg_DRETSEQ, 1, 3))

newDCRETRO.DRETEVT = Trim(txtfg_DRETEVT)
newDCRETRO.DRETNAT = Trim(txtfg_DRETNAT)
newDCRETRO.DRETREF = Trim(txtfg_DRETREF)
newDCRETRO.DRETCLI = Val(Mid$(txtfg_DRETCLI, 1, 7))
newDCRETRO.DRETDEV = Trim(Mid$(cbo_CDEV.Text, 1, 3))

' Montant tel qu'il est saisi : négatif ou positif
newDCRETRO.DRETMNT1 = CCur(Val(txtfg_DRETMNT1))
newDCRETRO.DRETMNT2 = CCur(Val(txtfg_DRETMNT2))
newDCRETRO.DRETCLR = Val(Mid$(txtfg_DRETCLR, 1, 7))
newDCRETRO.DRETCRTA = Val(Mid$(cbo_CRTA.Text, 1, 5))
newDCRETRO.DRETCTG = Val(txtfg_DRETCTG)

End Sub


Public Sub cmdDCRETRO_Update()

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

V = sqlDCRETRO_Update(newDCRETRO, xDCRETRO, cnADO)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

If Not IsNull(V) Then MsgBox V, vbCritical

End Sub

Public Sub cmdDCRETRO_Delete()

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

V = sqlDCRETRO_Delete(newDCRETRO, xDCRETRO, cnADO)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

If Not IsNull(V) Then MsgBox V, vbCritical

End Sub

Public Sub cmdDCRETRO_Insert()

Dim xSql As String, Nb As Integer
Dim wSWICLAOPR As String, wSWICLANUM As Long

xSql = "select * from " & paramIBM_Library_BODWH & ".DCRETRO " _
    & "where DRETVER = " & newDCRETRO.DRETVER _
    & " and DRETPER = " & newDCRETRO.DRETPER _
    & " and DRETETB = '" & newDCRETRO.DRETETB & "'" _
    & " and DRETAGE = '" & newDCRETRO.DRETAGE & "'" _
    & " and DRETSER = '" & newDCRETRO.DRETSER & "'" _
    & " and DRETSSE = '" & newDCRETRO.DRETSSE & "'" _
    & " and DRETOPE = '" & newDCRETRO.DRETOPE & "'" _
    & " and DRETNUM = " & newDCRETRO.DRETNUM _
    & " and DRETDTR = " & newDCRETRO.DRETDTR _
    & " and DRETSEQ = " & newDCRETRO.DRETSEQ

Set rsado = cnADO.Execute(xSql)
If Not rsado.EOF Then
    MsgBox "Enregistrement existant dans DCRETRO !", vbCritical, "fgDCRETRO.cmd_Creer_Insert"
    Exit Sub
End If

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

' wAmjMax = Période de saisie
newDCRETRO.DRETPER = wAmjMax
V = sqlDCRETRO_Insert(newDCRETRO, cnADO)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

        
If Not IsNull(V) Then

    MsgBox V, vbCritical
    cmdfra_Creer.Visible = True
    cmdfra_Modifier.Visible = False
    cmdfra_Supprimer.Visible = False
    txtfg_DRETVER.Enabled = False
    txtfg_DRETETB.Enabled = False
    txtfg_DRETAGE.Enabled = False
    cbo_CDEV.Enabled = True
    
    'txtfg_DRETCLIA.Enabled = True
    'txtfg_DRETCLIB.Enabled = True
    'cbo_CRTA.Enabled = True
    
    fraDCRETRO.Visible = True
    cmdSelect_Ok.Enabled = False
    cmdSelect_Ajouter.Enabled = False

Else
   
    ' Charger newDCRETRO dans arrDCRETRO puis
    ' Rafraichir l'affichage de la liste dans -fgSelect-
    ' Maintenir l'affichage de la ligne ajoutée sur la page en cours d'affichage
    
    arrDCRETRO_Nb = arrDCRETRO_Nb + 1
    arrDCRETRO(arrDCRETRO_Nb) = newDCRETRO
    xDCRETRO = newDCRETRO
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine arrDCRETRO_Nb

    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    fgSelect.TopRow = fgSelect.Row
     
    ' Maintenir l'affichage de la frame de gestion pour le mode -CREATION-
    xDCRETRO.DRETSTA = ""
    Call fraDCRETRO_Display

End If

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

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), BIA_DWH_Aut)
Form_Init

blnAuto = False

End Sub


Public Sub Form_Init()
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    If Not blnAuto Then MsgBox "paramétrage inconsistant", vbCritical, "frmBIA_DWH.param_init"
    Unload Me
Else
    lstErr.Clear
End If
If Not IsNull(paramSAA_Init) Then
    If Not blnAuto Then MsgBox "paramétrage inconsistant", vbCritical, "frmBIA_DWH.paramSAA_Init"
    Unload Me
Else
    lstErr.Clear
End If


blnControl = False
fgSelect_FormatString = fgSelect.FormatString

fgSelect.Enabled = True

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
fraSelect_Options.Enabled = True
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = False

blnControl = True

End Sub


Public Function param_Init()

param_Init = Null
Call lstErr_Clear(lstErr, cmdContext, "Param_Init"): DoEvents

fgSelect.Visible = False

Call DTPicker_Set(txtSelect_DRETPER, DSys)

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




Private Sub cbo_CDEV_GotFocus()
txt_GotFocus cbo_CDEV
End Sub


Private Sub cbo_CDEV_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub cbo_CDEV_LostFocus()
txt_LostFocus cbo_CDEV
End Sub


Private Sub cbo_CRTA_GotFocus()
txt_GotFocus cbo_CRTA
End Sub




Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdfra_Creer_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
cmdDCRETRO_Charger
cmdDCRETRO_Insert
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdfra_Modifier_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
cmdDCRETRO_Charger
cmdDCRETRO_Update
Me.Enabled = True: Me.MousePointer = 0
fraDCRETRO.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True

' Rafraichir la liste d'affichage -fgSelect-
arrDCRETRO(arrDCRETRO_Index) = newDCRETRO
xDCRETRO = newDCRETRO
fgSelect_DisplayLine arrDCRETRO_Index

End Sub


Private Sub cmdfra_Supprimer_Click()
Dim X As String
X = MsgBox("Confirmer la suppression ?", vbQuestion + vbOKCancel, "BIA_DWH : DCRETRO")

If X = vbOK Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    cmdDCRETRO_Charger
    cmdDCRETRO_Delete
    fraDCRETRO.Visible = False
    cmdSelect_Ok.Enabled = True
    cmdSelect_Ajouter.Enabled = True
    Me.Enabled = True: Me.MousePointer = 0
    
    fgSelect_TopRow_Memo = fgSelect.TopRow
    ' Rafraichir la liste sur -fgSelect- et positionner à la page en cours d'affichage
    cmdSelect_SQL
    If fgSelect.Rows > 1 Then fgSelect.TopRow = fgSelect_TopRow_Memo
Else
    fraDCRETRO.Visible = True
    cmdSelect_Ok.Enabled = False
    cmdSelect_Ajouter.Enabled = False
End If

End Sub


Private Sub cmdPrint_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case SSTab1.Tab
    Case 0:
           ' If fgSelect.Rows > 1 Then cmdPrint_Ok
                
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

currentAction = "cmdDCRETRO_SQL"
Call DTPicker_Control(txtSelect_DRETPER, wAmjMax)
xWhere_CRTA = " where DRGRPER = " & wAmjMax
xWhere = " where DRETPER = " & wAmjMax
  
DRTAGRP_CRTA_SQL xWhere_CRTA

xWhere_CDEV = ""
DDEVISE_CDEV_SQL xWhere_CDEV

arrDCRETRO_SQL xWhere
fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_Ajouter_Click()

Dim blnOk As Boolean, Nb As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_DWH_cmdSelect_Ajouter........"): DoEvents

xDCRETRO.DRETSTA = ""
Call fraDCRETRO_Display

Me.Enabled = True: Me.MousePointer = 0

End Sub



Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_DWH_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
fraDCRETRO.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True

If blnOk Then
    cmdSelect_Ok.Caption = "Nouveau mois"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = False
    cmdSelect_Ajouter.Visible = True
    cmdSelect_SQL
Else
    cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = True
    cmdSelect_Ajouter.Visible = False
End If
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_DWH_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

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
        fgSelect.Col = fgSelect_arrIndex:  arrDCRETRO_Index = CLng(fgSelect.Text)
        fgSelect.LeftCol = 0
        
        oldDCRETRO = arrDCRETRO(arrDCRETRO_Index)
        xDCRETRO = oldDCRETRO
        Call fraDCRETRO_Display
   End If
End If
fgSelect.LeftCol = 0
End Sub

Public Sub fraDCRETRO_DisplayLine()
Dim X As String

On Error Resume Next

 txtfg_DRETSTA = xDCRETRO.DRETSTA
 txtfg_DRETVER = xDCRETRO.DRETVER
 txtfg_DRETETB = xDCRETRO.DRETETB
 txtfg_DRETAGE = xDCRETRO.DRETAGE
 txtfg_DRETSER = xDCRETRO.DRETSER
 txtfg_DRETSSE = xDCRETRO.DRETSSE
 txtfg_DRETOPE = xDCRETRO.DRETOPE
 txtfg_DRETNUM = xDCRETRO.DRETNUM
 
 wAmjDTR = xDCRETRO.DRETDTR
 Call DTPicker_Set(txtfg_DRETDTR, wAmjDTR)
 txtfg_DRETSEQ = xDCRETRO.DRETSEQ

 txtfg_DRETEVT = xDCRETRO.DRETEVT
 txtfg_DRETNAT = xDCRETRO.DRETNAT
 txtfg_DRETREF = xDCRETRO.DRETREF
 txtfg_DRETCLI = xDCRETRO.DRETCLI
 Call cbo_Scan(xDCRETRO.DRETDEV, cbo_CDEV)
 txtfg_DRETMNT1 = cur_P(xDCRETRO.DRETMNT1)
 txtfg_DRETMNT2 = cur_P(xDCRETRO.DRETMNT2)
 txtfg_DRETCLR = xDCRETRO.DRETCLR
 
 ' Le code renta du DCRETRO peut être en 5 ou 4 ou 3 numériques rempli
 X = Format$(xDCRETRO.DRETCRTA, "00000")
 If Mid$(X, 1, 1) = "0" Then
    X = Format$(X, "0000")
    If Mid$(X, 1, 1) = "0" Then
        X = Format$(X, "000")
    End If
 End If
 Call cbo_Scan(X, cbo_CRTA)
 
 txtfg_DRETCTG = xDCRETRO.DRETCTG

End Sub

Private Sub fraDCRETRO_Display()
Dim V
Dim X As String, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim I As Long

On Error GoTo Error_Handler

fraDCRETRO.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True
cmdfra_Creer.Visible = False
cmdfra_Modifier.Visible = False
cmdfra_Supprimer.Visible = False

currentAction = "fraDCRETRO_Display"

If Trim(xDCRETRO.DRETSTA) <> "" Then   ' Autres que Création
   
    fraDCRETRO_DisplayLine
    
    cmdfra_Creer.Visible = False
    cmdfra_Modifier.Visible = True
    cmdfra_Supprimer.Visible = True
    txtfg_DRETVER.Enabled = False
    txtfg_DRETETB.Enabled = False
    txtfg_DRETAGE.Enabled = False
    txtfg_DRETSER.Enabled = False
    txtfg_DRETSSE.Enabled = False
    txtfg_DRETOPE.Enabled = False
    txtfg_DRETNUM.Enabled = False
    txtfg_DRETDTR.Enabled = False
    txtfg_DRETSEQ.Enabled = False

Else  ' Création
   
    txtfg_DRETSTA = "W"
    txtfg_DRETVER = 1
    txtfg_DRETETB = "01"
    txtfg_DRETAGE = "01"
    txtfg_DRETSER = "00"
    txtfg_DRETSSE = "00"
    txtfg_DRETOPE = ""
    txtfg_DRETNUM = ""
    Call DTPicker_Set(txtfg_DRETDTR, DSys)
    txtfg_DRETSEQ = ""
    txtfg_DRETEVT = ""
    txtfg_DRETNAT = ""
    txtfg_DRETREF = ""
    txtfg_DRETCLI = ""
    txtfg_DRETMNT1 = ""
    txtfg_DRETMNT2 = ""
    txtfg_DRETCLR = ""
    txtfg_DRETCTG = 1

    cmdfra_Creer.Visible = True
    cmdfra_Modifier.Visible = False
    cmdfra_Supprimer.Visible = False
    txtfg_DRETVER.Enabled = False
    txtfg_DRETETB.Enabled = False
    txtfg_DRETAGE.Enabled = False
    
    txtfg_DRETSER.Enabled = True
    txtfg_DRETSSE.Enabled = True
    txtfg_DRETOPE.Enabled = True
    txtfg_DRETNUM.Enabled = True
    txtfg_DRETDTR.Enabled = True
    txtfg_DRETSEQ.Enabled = True

End If

fraDCRETRO.Visible = True
cmdSelect_Ok.Enabled = False
cmdSelect_Ajouter.Enabled = False

Call lstErr_AddItem(lstErr, cmdContext, "Lignes : "): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 1
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

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
If fraDCRETRO.Visible Then
    fraDCRETRO.Visible = False
    cmdSelect_Ok.Enabled = True
    cmdSelect_Ajouter.Enabled = True
    Exit Sub
End If
    
If SSTab1.Tab = 0 Then
        Unload Me
    Exit Sub
Else
    SSTab1.Tab = SSTab1.Tab - 1
End If

End Sub

Public Sub cmdContext_Return()
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

Set cnADO = New ADODB.Connection
cnADO.CursorLocation = adUseClient
cnADO.Open paramODBC_DSN_SAB

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






Public Sub blnTransaction_Set()
If Not blnTransaction Then
    blnTransaction = True
    Set rsADO_Update = cnADO.Execute("SET TRANSACTION ISOLATION LEVEL READ COMMITTED")

End If

End Sub























Private Sub txtfg_DRETAGE_GotFocus()
txt_GotFocus txtfg_DRETAGE
End Sub


Private Sub txtfg_DRETAGE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DRETCLI_GotFocus()
txt_GotFocus txtfg_DRETCLI
End Sub


Private Sub txtfg_DRETCLI_LostFocus()
txt_LostFocus txtfg_DRETCLI

End Sub


Private Sub txtfg_DRETCLR_GotFocus()
txt_GotFocus txtfg_DRETCLR
End Sub

Private Sub txtfg_DRETCLR_LostFocus()
txt_LostFocus txtfg_DRETCLR
End Sub


Private Sub txtfg_DRETCTG_GotFocus()
txt_GotFocus txtfg_DRETCTG

End Sub


Private Sub txtfg_DRETETB_GotFocus()
txt_GotFocus txtfg_DRETETB

End Sub


Private Sub txtfg_DRETETB_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtfg_DRETEVT_GotFocus()
txt_GotFocus txtfg_DRETEVT

End Sub


Private Sub txtfg_DRETEVT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DRETMNT1_KeyPress(KeyAscii As Integer)
num_KeyAsciiD KeyAscii, txtfg_DRETMNT1
End Sub


Private Sub txtfg_DRETMNT1_LostFocus()
txtfg_DRETMNT1 = cur_P(CCur(Fix(Val(txtfg_DRETMNT1) * 100) / 100))
End Sub


Private Sub txtfg_DRETMNT2_KeyPress(KeyAscii As Integer)
num_KeyAsciiD KeyAscii, txtfg_DRETMNT2
End Sub


Private Sub txtfg_DRETMNT2_LostFocus()
txtfg_DRETMNT2 = cur_P(CCur(Fix(Val(txtfg_DRETMNT2) * 100) / 100))
End Sub


Private Sub txtfg_DRETNAT_GotFocus()
txt_GotFocus txtfg_DRETNAT

End Sub


Private Sub txtfg_DRETNAT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DRETNUM_GotFocus()
txt_GotFocus txtfg_DRETNUM

End Sub


Private Sub txtfg_DRETNUM_LostFocus()
txt_LostFocus txtfg_DRETNUM

End Sub


Private Sub txtfg_DRETOPE_GotFocus()
txt_GotFocus txtfg_DRETOPE

End Sub


Private Sub txtfg_DRETOPE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DRETREF_GotFocus()
txt_GotFocus txtfg_DRETREF

End Sub


Private Sub txtfg_DRETREF_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DRETREF_LostFocus()
txt_LostFocus txtfg_DRETREF

End Sub

Private Sub txtfg_DRETSEQ_GotFocus()
txt_GotFocus txtfg_DRETSEQ

End Sub


Private Sub txtfg_DRETSER_GotFocus()
txt_GotFocus txtfg_DRETSER

End Sub

Private Sub txtfg_DRETSER_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DRETSSE_GotFocus()
txt_GotFocus txtfg_DRETSSE

End Sub

Private Sub txtfg_DRETSSE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DRETSTA_GotFocus()
txt_GotFocus txtfg_DRETSTA
End Sub

Private Sub txtfg_DRETSTA_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DRETSTA_LostFocus()
txt_LostFocus txtfg_DRETSTA

End Sub





Private Sub txtfg_DRETVER_GotFocus()
txt_GotFocus txtfg_DRETVER
End Sub





