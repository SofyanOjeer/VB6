VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDRENTACH 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA_DWH"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13560
   Icon            =   "DRENTACH.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9150
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
      TabCaption(0)   =   "Gestion sur les mois traités"
      TabPicture(0)   =   "DRENTACH.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "....."
      TabPicture(1)   =   "DRENTACH.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13290
         Begin VB.CommandButton cmdSelect_Ajouter 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Ajouter"
            Height          =   285
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   720
            Width           =   1335
         End
         Begin VB.Frame fraDRENTACH 
            BackColor       =   &H0080C0FF&
            Caption         =   "Gestion du fichier des charges - DRENTACH -"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   5775
            Left            =   6840
            TabIndex        =   11
            Top             =   1680
            Visible         =   0   'False
            Width           =   6015
            Begin VB.ComboBox cbo_CRTA 
               Height          =   315
               Left            =   2160
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   2040
               Width           =   2655
            End
            Begin VB.CommandButton cmdfra_Creer 
               BackColor       =   &H00FFC0FF&
               Caption         =   "Créer"
               Height          =   495
               Left            =   5040
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   4920
               Width           =   855
            End
            Begin VB.CommandButton cmdfra_Modifier 
               BackColor       =   &H00FF80FF&
               Caption         =   "Modifier"
               Height          =   495
               Left            =   3960
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   4920
               Width           =   855
            End
            Begin VB.CommandButton cmdfra_Supprimer 
               BackColor       =   &H000000FF&
               Caption         =   "Supprimer"
               Height          =   495
               Left            =   2760
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   4920
               Width           =   855
            End
            Begin VB.TextBox txtfg_DRCHMMRB 
               Height          =   300
               Left            =   2160
               TabIndex        =   20
               Text            =   "Montant marge"
               Top             =   3000
               Width           =   1695
            End
            Begin VB.TextBox txtfg_DRCHCTR 
               Height          =   300
               Left            =   2160
               TabIndex        =   21
               Text            =   "Comptage"
               Top             =   3480
               Width           =   615
            End
            Begin VB.TextBox txtfg_DRCHCGRP 
               Height          =   300
               Left            =   2160
               TabIndex        =   19
               Text            =   "Code GRP renta"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtfg_DRCHCLIB 
               Enabled         =   0   'False
               Height          =   300
               Left            =   3000
               TabIndex        =   17
               Text            =   "No matricule"
               Top             =   1560
               Width           =   855
            End
            Begin VB.TextBox txtfg_DRCHCLIA 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2160
               TabIndex        =   16
               Text            =   "Blanc / T"
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox txtfg_DRCHETA 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2160
               TabIndex        =   15
               Text            =   "Etablissement"
               Top             =   1080
               Width           =   615
            End
            Begin VB.TextBox txtfg_DRCHVER 
               Enabled         =   0   'False
               Height          =   300
               Left            =   3000
               TabIndex        =   14
               Text            =   "Version"
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox txtfg_DRCHSTA 
               Height          =   300
               Left            =   2160
               TabIndex        =   13
               Text            =   "Statut"
               Top             =   600
               Width           =   615
            End
            Begin VB.Label lblfg_DRCHCTR 
               BackColor       =   &H0080C0FF&
               Caption         =   "Comptage"
               Height          =   300
               Left            =   240
               TabIndex        =   27
               Top             =   3480
               Width           =   1815
            End
            Begin VB.Label lblfg_DRCHMMRB 
               BackColor       =   &H0080C0FF&
               Caption         =   "Montant marge renta - BASE"
               Height          =   375
               Left            =   240
               TabIndex        =   26
               Top             =   3000
               Width           =   1815
            End
            Begin VB.Label lblfg_DRCHCGRP 
               BackColor       =   &H0080C0FF&
               Caption         =   "Code regroupement renta"
               Height          =   300
               Left            =   240
               TabIndex        =   25
               Top             =   2520
               Width           =   1815
            End
            Begin VB.Label lblfg_DRCHCRTA 
               BackColor       =   &H0080C0FF&
               Caption         =   "Code renta"
               Height          =   300
               Left            =   240
               TabIndex        =   24
               Top             =   2040
               Width           =   1815
            End
            Begin VB.Label lblfg_DRCHCLIA_CLIB 
               BackColor       =   &H0080C0FF&
               Caption         =   "Blanc / T / No matricule"
               Height          =   300
               Left            =   240
               TabIndex        =   23
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label lblfg_DRCHETA 
               BackColor       =   &H0080C0FF&
               Caption         =   "Code établissement"
               Height          =   300
               Left            =   240
               TabIndex        =   22
               Top             =   1080
               Width           =   1815
            End
            Begin VB.Label lblfg_DRCHSTA_VER 
               BackColor       =   &H0080C0FF&
               Caption         =   "Statut / Version"
               Height          =   300
               Left            =   240
               TabIndex        =   12
               Top             =   600
               Width           =   1815
            End
         End
         Begin VB.Frame fraSelect_Options 
            Height          =   1005
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   11355
            Begin MSComCtl2.DTPicker txtSelect_DRCHPER 
               Height          =   300
               Left            =   1920
               TabIndex        =   8
               Top             =   240
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   54198275
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblSelect_DRCHPER 
               Caption         =   "Période de traitement"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   315
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6825
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   12840
            _ExtentX        =   22648
            _ExtentY        =   12039
            _Version        =   393216
            Rows            =   1
            Cols            =   11
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
            FormatString    =   "Statut|<Version|>Période   |Etab |  /<Matricule |>Code renta|>Code regroupement|> Montant marge renta -Base |>Comptage ||"
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
      Picture         =   "DRENTACH.frx":047A
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
End
Attribute VB_Name = "frmDRENTACH"
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

Dim cnAdo As New ADODB.Connection, rsAdo As New ADODB.Recordset, errADO As ADODB.Error
Dim blnTransaction As Boolean

'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
Dim xDRENTACH As typeDRENTACH, newDRENTACH As typeDRENTACH, oldDRENTACH As typeDRENTACH
Dim arrDRENTACH() As typeDRENTACH, arrDRENTACH_Nb As Long, arrDRENTACH_Max As Long, arrDRENTACH_Index As Long
Dim xDRTAGRP As typeDRTAGRP
Dim xDRLCRTA As typeDRLCRTA, arrDRLCRTA(100) As typeDRLCRTA, arrDRLCRTA_Nb As Long
Dim xWhere_CRTA As String, xWhere_CGRP As String

'______________________________________________________________________

Dim fraDRENTACH_FormatString As String, fraDRENTACH_K As Integer
Dim fraDRENTACH_RowDisplay As Integer, fraDRENTACH_RowClick As Integer, fraDRENTACH_ColClick As Integer
Dim fraDRENTACH_ColorClick As Long, fraDRENTACH_ColorDisplay As Long
Dim fraDRENTACH_Sort1 As Integer, fraDRENTACH_Sort2 As Integer
Dim fraDRENTACH_SortAD As Integer, fraDRENTACH_Sort1_Old As Integer
Dim fraDRENTACH_arrIndex As Integer
Dim blnfraDRENTACH_DisplayLine As Boolean

Dim meDRENTACH_Status As typeDRENTACH, oldDRENTACH_Status As typeDRENTACH

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
    
For I = 1 To arrDRENTACH_Nb
         
    xDRENTACH = arrDRENTACH(I)
    If xDRENTACH.DRCHSTA <> "" Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
    End If
Next I

fgSelect.Visible = True
fraDRENTACH.Visible = False
cmdSelect_Ok.Enabled = True
'cmdSelect_Ajouter.Enabled = True

Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrDRENTACH_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
    
End Sub

Private Sub arrDRENTACH_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrDRENTACH(101)
arrDRENTACH_Max = 100: arrDRENTACH_Nb = 0

Set rsAdo = Nothing

xSQL = "select * from " & paramIBM_Library_BODWH & ".DRENTACH " & xWhere & " order by DRCHCLIA, DRCHCLIB, DRCHCRTA"
Set rsAdo = cnAdo.Execute(xSQL)

Do While Not rsAdo.EOF
    V = srvDRENTACH_GetBuffer_ODBC(rsAdo, xDRENTACH)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDRENTACH.fgselect_Display"
        '' Exit Sub
     Else
         arrDRENTACH_Nb = arrDRENTACH_Nb + 1
         If arrDRENTACH_Nb > arrDRENTACH_Max Then
             arrDRENTACH_Max = arrDRENTACH_Max + 50
             ReDim Preserve arrDRENTACH(arrDRENTACH_Max)
         End If
         
         arrDRENTACH(arrDRENTACH_Nb) = xDRENTACH
    End If
    rsAdo.MoveNext
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
Dim X As String, xSQL As String
Dim I As Integer

On Error GoTo Error_Handler

cbo_CRTA.Clear

Set rsAdo = Nothing

xSQL = "select * from " & paramIBM_Library_BODWH & ".DRTAGRP " & xWhere_CRTA & " order by DRGRCRTA"
Set rsAdo = cnAdo.Execute(xSQL)

Do While Not rsAdo.EOF
    V = srvDRTAGRP_GetBuffer_ODBC(rsAdo, xDRTAGRP)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDRENTACH.DRTAGRP_CRTA_SQL"
        '' Exit Sub
     Else
         For I = 0 To arrDRLCRTA_Nb - 1  ' arrDRLCRTA_Nb chargé dans DRLCRTA_CHARGES_SQL
            xDRLCRTA = arrDRLCRTA(I)
            If xDRLCRTA.DRRTACRTA = xDRTAGRP.DRGRCRTA Then
               cbo_CRTA.AddItem xDRTAGRP.DRGRCRTA & "  " & xDRTAGRP.DRGRLIB
            End If
         Next I
    End If
    rsAdo.MoveNext
Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub DRLCRTA_CHARGES_SQL()

Dim V
Dim X As String, xSQL As String

On Error GoTo Error_Handler

arrDRLCRTA_Nb = 0

Set rsAdo = Nothing

xSQL = "select * from " & paramIBM_Library_BODWH & ".DRLCRTA order by DRRTACRTA"
Set rsAdo = cnAdo.Execute(xSQL)

Do While Not rsAdo.EOF
    V = srvDRLCRTA_GetBuffer_ODBC(rsAdo, xDRLCRTA)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDRENTACH.DRLCRTA_CHARGES_SQL"
        '' Exit Sub
     Else
        If xDRLCRTA.DRRTANAT = "D" And xDRLCRTA.DRRTACRTA > 1000 Then
            arrDRLCRTA(arrDRLCRTA_Nb) = xDRLCRTA
            ' !!! Le tableau arrDRLCRTA fixé à 100 postes
            arrDRLCRTA_Nb = arrDRLCRTA_Nb + 1
        End If
    End If
    rsAdo.MoveNext
Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub DRTAGRP_CGRP_SQL(xWhere_CGRP As String)
Dim V
Dim wLecture As Boolean
Dim X As String, xSQL As String
On Error GoTo Error_Handler

wLecture = False
Set rsAdo = Nothing

xSQL = "select *from " & paramIBM_Library_BODWH & ".DRTAGRP where DRGRCRTA = " & Val(cbo_CRTA.Text) & xWhere_CGRP
Set rsAdo = cnAdo.Execute(xSQL)

Do While Not rsAdo.EOF
    V = srvDRTAGRP_GetBuffer_ODBC(rsAdo, xDRTAGRP)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmDRENTACH.DRTAGRP_CGRP_SQL"
         Exit Sub
     Else
         txtfg_DRCHCGRP = xDRTAGRP.DRGRCGRP
         wLecture = True
     End If
    rsAdo.MoveNext
Loop

If wLecture = False Then
    X = MsgBox("Ressaisir le code renta inconnu !", vbCritical, "BIA_DWH : DRENTACH")
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Public Sub fgSelect_DisplayLine(lIndex As Long)

On Error Resume Next

fgSelect.Col = 0: fgSelect.Text = xDRENTACH.DRCHSTA
fgSelect.Col = 1: fgSelect.Text = xDRENTACH.DRCHVER
fgSelect.Col = 2: fgSelect.Text = dateImp10(xDRENTACH.DRCHPER)
fgSelect.Col = 3: fgSelect.Text = xDRENTACH.DRCHETA
fgSelect.Col = 4: fgSelect.Text = xDRENTACH.DRCHCLIA & " / " & xDRENTACH.DRCHCLIB

fgSelect.Col = 5: fgSelect.Text = xDRENTACH.DRCHCRTA
fgSelect.Col = 6: fgSelect.Text = xDRENTACH.DRCHCGRP
fgSelect.Col = 7: fgSelect.Text = Format$(xDRENTACH.DRCHMMRB, "### ### ### ##0.00")

If xDRENTACH.DRCHMMRB = 0 Then
    fgSelect.Text = ""
Else
    ' Mettre en rouge les montants débiteurs
    If xDRENTACH.DRCHMMRB >= 0 Then
        fgSelect.CellForeColor = vbBlue
    Else
        fgSelect.CellForeColor = vbRed
    End If
End If

fgSelect.Col = 8: fgSelect.Text = xDRENTACH.DRCHCTR

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

Public Sub cmdDRENTACH_Charger()

srvDRENTACH_Init newDRENTACH

newDRENTACH.DRCHSTA = Trim(txtfg_DRCHSTA)
newDRENTACH.DRCHVER = Val(txtfg_DRCHVER)

newDRENTACH.DRCHPER = wAmjMax   '''xDRENTACH.DRCHPER
newDRENTACH.DRCHETA = Trim(txtfg_DRCHETA)

newDRENTACH.DRCHCLIA = Trim(txtfg_DRCHCLIA)
newDRENTACH.DRCHCLIB = Val(Mid$(txtfg_DRCHCLIB, 1, 7))
newDRENTACH.DRCHCRTA = Val(Mid$(cbo_CRTA.Text, 1, 5))
newDRENTACH.DRCHCGRP = Val(txtfg_DRCHCGRP)

' DRENTACH est destiné aux charges DONC le montant est systématiquement signé NEGATIF
' Pour EVITER une saisie de montant négatif SYSTEMATIQUEMENT
newDRENTACH.DRCHMMRB = CCur(Val(txtfg_DRCHMMRB))
newDRENTACH.DRCHMMRB = Abs(newDRENTACH.DRCHMMRB) * -1

newDRENTACH.DRCHCTR = Val(txtfg_DRCHCTR)

End Sub


Public Sub cmdDRENTACH_Update()

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

V = sqlDRENTACH_Update(newDRENTACH, xDRENTACH, cnAdo)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

If Not IsNull(V) Then MsgBox V, vbCritical

End Sub

Public Sub cmdDRENTACH_Delete()

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

V = sqlDRENTACH_Delete(newDRENTACH, xDRENTACH, cnAdo)

'If Not IsNull(V) Then
'    xSql = "Rollback"
'Else
'    xSql = "Commit"
'End If

'Set rsADO_Update = cnAdo.Execute(xSql, Nb)
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

If Not IsNull(V) Then MsgBox V, vbCritical

End Sub

Public Sub cmdDRENTACH_Insert()

Dim xSQL As String, Nb As Integer
Dim wSWICLAOPR As String, wSWICLANUM As Long

xSQL = "select * from " & paramIBM_Library_BODWH & ".DRENTACH " _
    & "where DRCHVER = " & newDRENTACH.DRCHVER _
    & " and DRCHPER = " & newDRENTACH.DRCHPER _
    & " and DRCHETA = '" & newDRENTACH.DRCHETA & "'" _
    & " and DRCHCLIA = '" & newDRENTACH.DRCHCLIA & "'" _
    & " and DRCHCLIB = " & newDRENTACH.DRCHCLIB _
    & " and DRCHCRTA = " & newDRENTACH.DRCHCRTA

Set rsAdo = cnAdo.Execute(xSQL)
If Not rsAdo.EOF Then
    MsgBox "Enregistrement existant dans DRENTACH !", vbCritical, "fgDRENTACH.cmd_Creer_Insert"
    Exit Sub
End If

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'blnTransaction_Set

' wAmjMax = Période de saisie
newDRENTACH.DRCHPER = wAmjMax
V = sqlDRENTACH_Insert(newDRENTACH, cnAdo)

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
    txtfg_DRCHVER.Enabled = False
    txtfg_DRCHETA.Enabled = False
    txtfg_DRCHCLIA.Enabled = True
    txtfg_DRCHCLIB.Enabled = True
    cbo_CRTA.Enabled = True
    
    fraDRENTACH.Visible = True
    cmdSelect_Ok.Enabled = False
   ' cmdSelect_Ajouter.Enabled = False

Else
   
    ' Charger newDRENTACH dans arrDRENTACH puis
    ' Rafraichir l'affichage de la liste dans -fgSelect-
    ' Maintenir l'affichage de la ligne ajoutée sur la page en cours d'affichage
    
    arrDRENTACH_Nb = arrDRENTACH_Nb + 1
    arrDRENTACH(arrDRENTACH_Nb) = newDRENTACH
    xDRENTACH = newDRENTACH
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine arrDRENTACH_Nb

    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    fgSelect.TopRow = fgSelect.Row
   
    ' Maintenir l'affichage de la frame de gestion pour le mode -CREATION-
    xDRENTACH.DRCHSTA = ""
    Call fraDRENTACH_Display

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

Call DTPicker_Set(txtSelect_DRCHPER, DSys)

' Lecture 1 seule fois des codes renta de CHARGES
DRLCRTA_CHARGES_SQL

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




Private Sub cbo_CRTA_GotFocus()
txt_GotFocus cbo_CRTA
End Sub


Private Sub cbo_CRTA_LostFocus()
txt_LostFocus cbo_CRTA
DRTAGRP_CGRP_SQL xWhere_CGRP
End Sub


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdfra_Creer_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
cmdDRENTACH_Charger
cmdDRENTACH_Insert
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdfra_Modifier_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
cmdDRENTACH_Charger
cmdDRENTACH_Update
fraDRENTACH.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = True

' Rafraichir la liste d'affichage -fgSelect-
arrDRENTACH(arrDRENTACH_Index) = newDRENTACH
xDRENTACH = newDRENTACH
fgSelect_DisplayLine arrDRENTACH_Index
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdfra_Supprimer_Click()
Dim X As String
X = MsgBox("Confirmer la suppression ?", vbQuestion + vbOKCancel, "BIA_DWH : DRENTACH")

If X = vbOK Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    cmdDRENTACH_Charger
    cmdDRENTACH_Delete
    fraDRENTACH.Visible = False
    cmdSelect_Ok.Enabled = True
    cmdSelect_Ajouter.Enabled = True
    
    fgSelect_TopRow_Memo = fgSelect.TopRow
    ' Rafraichir la liste sur -fgSelect- et positionner à la page en cours d'affichage
    cmdSelect_SQL
    If fgSelect.Rows > 1 Then fgSelect.TopRow = fgSelect_TopRow_Memo
    Me.Enabled = True: Me.MousePointer = 0
Else
    fraDRENTACH.Visible = True
    cmdSelect_Ok.Enabled = False
    'cmdSelect_Ajouter.Enabled = False
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

currentAction = "cmdDRENTACH_SQL"
Call DTPicker_Control(txtSelect_DRCHPER, wAmjMax)
xWhere_CRTA = " where DRGRPER = " & wAmjMax
xWhere_CGRP = " and DRGRPER = " & wAmjMax
xWhere = " where DRCHPER = " & wAmjMax
  
DRTAGRP_CRTA_SQL xWhere_CRTA

arrDRENTACH_SQL xWhere
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

xDRENTACH.DRCHSTA = ""
Call fraDRENTACH_Display

Me.Enabled = True: Me.MousePointer = 0

End Sub



Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_DWH_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
fraDRENTACH.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = False

If blnOk Then
    cmdSelect_Ok.Caption = "Nouveau mois"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = False
    cmdSelect_Ajouter.Enabled = True
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


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
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
        fgSelect.Col = fgSelect_arrIndex:  arrDRENTACH_Index = CLng(fgSelect.Text)
        fgSelect.LeftCol = 0
        
        oldDRENTACH = arrDRENTACH(arrDRENTACH_Index)
        xDRENTACH = oldDRENTACH
        Call fraDRENTACH_Display
   End If
End If
fgSelect.LeftCol = 0
End Sub

Public Sub fraDRENTACH_DisplayLine()
Dim X As String

On Error Resume Next

 txtfg_DRCHSTA = xDRENTACH.DRCHSTA
 txtfg_DRCHVER = xDRENTACH.DRCHVER
 txtfg_DRCHETA = xDRENTACH.DRCHETA
 txtfg_DRCHCLIA = xDRENTACH.DRCHCLIA
 txtfg_DRCHCLIB = xDRENTACH.DRCHCLIB
 
 ' Code renta charges n'est que sur 4 numériques
 X = Format$(xDRENTACH.DRCHCRTA, "0000")
 Call cbo_Scan(X, cbo_CRTA)
 
 txtfg_DRCHCGRP = xDRENTACH.DRCHCGRP
 txtfg_DRCHMMRB = cur_P(xDRENTACH.DRCHMMRB)
 txtfg_DRCHCTR = xDRENTACH.DRCHCTR

End Sub

Private Sub fraDRENTACH_Display()
Dim V
Dim X As String, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim I As Long

On Error GoTo Error_Handler

fraDRENTACH.Visible = False
cmdSelect_Ok.Enabled = True
cmdSelect_Ajouter.Enabled = False
cmdfra_Creer.Visible = False
cmdfra_Modifier.Visible = False
cmdfra_Supprimer.Visible = False

currentAction = "fraDRENTACH_Display"

If Trim(xDRENTACH.DRCHSTA) <> "" Then   ' Autres que Création
   
    fraDRENTACH_DisplayLine
    
    cmdfra_Creer.Visible = False
    cmdfra_Modifier.Visible = True
    cmdfra_Supprimer.Visible = True
    txtfg_DRCHVER.Enabled = False
    txtfg_DRCHETA.Enabled = False
    txtfg_DRCHCLIA.Enabled = False
    txtfg_DRCHCLIB.Enabled = False
    cbo_CRTA.Enabled = False

Else  ' Création
   
    txtfg_DRCHSTA = "P"
    txtfg_DRCHVER = 1
    txtfg_DRCHETA = "01"
    txtfg_DRCHCLIA = ""
    txtfg_DRCHCLIB = ""
    txtfg_DRCHCGRP = ""
    txtfg_DRCHMMRB = ""
    txtfg_DRCHCTR = 1

    cmdfra_Creer.Visible = True
    cmdfra_Modifier.Visible = False
    cmdfra_Supprimer.Visible = False
    txtfg_DRCHVER.Enabled = False
    txtfg_DRCHETA.Enabled = False
    txtfg_DRCHCLIA.Enabled = True
    txtfg_DRCHCLIB.Enabled = True
    cbo_CRTA.Enabled = True

End If

fraDRENTACH.Visible = True
cmdSelect_Ok.Enabled = False
'cmdSelect_Ajouter.Enabled = False

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
    cnAdo.Close
    Set cnAdo = Nothing

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
If fraDRENTACH.Visible Then
    fraDRENTACH.Visible = False
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
'    SendKeys "{TAB}"
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

Set cnAdo = New ADODB.Connection
cnAdo.CursorLocation = adUseClient
cnAdo.Open paramODBC_DSN_SAB

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
    Set rsADO_Update = cnAdo.Execute("SET TRANSACTION ISOLATION LEVEL READ COMMITTED")

End If

End Sub

Private Sub txtfg_DRCHCGRP_GotFocus()
txt_GotFocus txtfg_DRCHCGRP
DRTAGRP_CGRP_SQL xWhere_CGRP
End Sub

Private Sub txtfg_DRCHCGRP_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtfg_DRCHCGRP_LostFocus()
txt_LostFocus txtfg_DRCHCGRP

End Sub

Private Sub txtfg_DRCHCLIA_GotFocus()
txt_GotFocus txtfg_DRCHCLIA
End Sub

Private Sub txtfg_DRCHCLIA_LostFocus()
txt_LostFocus txtfg_DRCHCLIA

End Sub


Private Sub txtfg_DRCHCLIB_GotFocus()
txt_GotFocus txtfg_DRCHCLIB
End Sub




Private Sub txtfg_DRCHCLIB_LostFocus()
txt_LostFocus txtfg_DRCHCLIB

End Sub


Private Sub txtfg_DRCHCTR_GotFocus()
txt_GotFocus txtfg_DRCHCTR
End Sub


Private Sub txtfg_DRCHCTR_LostFocus()
txt_LostFocus txtfg_DRCHCTR

End Sub


Private Sub txtfg_DRCHETA_GotFocus()
txt_GotFocus txtfg_DRCHETA
End Sub

Private Sub txtfg_DRCHETA_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtfg_DRCHETA_LostFocus()
txt_LostFocus txtfg_DRCHETA

End Sub

Private Sub txtfg_DRCHMMRB_KeyPress(KeyAscii As Integer)
num_KeyAsciiD KeyAscii, txtfg_DRCHMMRB
End Sub


Private Sub txtfg_DRCHMMRB_LostFocus()
'txtfg_DRCHMMRB = cur_P((CCur(Fix(Val(txtfg_DRCHMMRB) * 100) + 0.00500001) / 100))
txtfg_DRCHMMRB = cur_P(Val(txtfg_DRCHMMRB))
End Sub

Private Sub txtfg_DRCHSTA_GotFocus()
txt_GotFocus txtfg_DRCHSTA
End Sub

Private Sub txtfg_DRCHSTA_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtfg_DRCHSTA_LostFocus()
txt_LostFocus txtfg_DRCHSTA

End Sub


Private Sub txtfg_DRCHVER_GotFocus()
txt_GotFocus txtfg_DRCHVER
End Sub


