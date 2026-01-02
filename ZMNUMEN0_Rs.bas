Attribute VB_Name = "rsZMNUMEN0"
Option Explicit

Type typeZMNUMEN0
    
    MNUMENETB       As Integer                        ' ETABLISSEMENT
    MNUMENREF       As Long                           ' REFERENCE LOT
    MNUMENGRP       As String * 10                    ' GROUPE MENU
    MNUMENPRE       As Long                           ' CODE OPTION PRECED
    MNUMENORD       As Long                           ' ORDRE DANS MENU
    MNUMENCOD       As Long                           ' CODE OPTION
    MNUMENOIA       As String * 1                     ' INTER-AGENCE
    MNUMENJOQ       As String * 10                    ' FILE ATTENT.BATCH
    
    Method          As String
    Niveau          As Long
    Hierarchie      As String
End Type

Public Sub arrZMNUMEN0_Load(lZMNUHLB0 As typeZMNUHLB0, arrZMNUMEN0() As typeZMNUMEN0, arrZMNUOPT0() As typeZMNUOPT0)
Dim V, xSQL As String
Dim I As Integer, K As Integer
Dim wNb As Integer
Dim blnOk As Boolean
Dim kHierarchie As Integer
On Error GoTo Error_Handler

Set rsSab = Nothing
'-------------------------------------------------------
App_Debug = "arrZMNUMEN0_Load : " & lZMNUHLB0.MNUHLBNOM
'-------------------------------------------------------
ReDim arrZMNUMEN0(2000), arrZMNUOPT0(2000)
wNb = 0
xSQL = "select * from " & paramIBM_Library_SAB & ".ZMNUMEN0 , " & paramIBM_Library_SAB & ".ZMNUOPT0" _
     & " where MNUMENGRP = '" & lZMNUHLB0.MNUHLBNOM & "'" _
     & " and   MNUMENREF =" & lZMNUHLB0.MNUHLBREF _
     & " and   MNUMENETB =" & lZMNUHLB0.MNUHLBETB _
     & " and   MNUMENCOD = MNUOPTCOD" _
     & " order by MNUMENPRE, MNUMENORD"
     
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    If wNb = UBound(arrZMNUMEN0) Then
        ReDim Preserve arrZMNUMEN0(wNb + 500)
        ReDim Preserve arrZMNUOPT0(wNb + 500)
    End If
    wNb = wNb + 1
    Call rsZMNUMEN0_GetBuffer(rsSab, arrZMNUMEN0(wNb))
    Call rsZMNUOPT0_GetBuffer(rsSab, arrZMNUOPT0(wNb))
    rsSab.MoveNext
Loop

K = 0: arrZMNUMEN0(0).MNUMENCOD = 0: arrZMNUMEN0(0).Hierarchie = "": arrZMNUMEN0(0).Niveau = 0
kHierarchie = 0
Do
    blnOk = True
    kHierarchie = kHierarchie + 1
    For I = 1 To wNb
        If arrZMNUMEN0(I).Hierarchie = "" Then
            If arrZMNUMEN0(I).MNUMENPRE <> arrZMNUMEN0(K).MNUMENCOD Then
                For K = 1 To wNb
                    If arrZMNUMEN0(I).MNUMENPRE = arrZMNUMEN0(K).MNUMENCOD Then Exit For
                Next K
            End If
            If arrZMNUMEN0(I).MNUMENPRE <> 0 And arrZMNUMEN0(I).MNUMENPRE <> 9999999 And arrZMNUMEN0(K).Hierarchie = "" Then
                If kHierarchie < 5 Then
                    blnOk = False
                Else
                    arrZMNUMEN0(I).Hierarchie = Format(I, "00000")
                    arrZMNUMEN0(I).Niveau = -1
                   '' MsgBox arrZMNUMEN0(I).MNUMENCOD & " : " & arrZMNUOPT0(I).MNUOPTLIB, vbInformation, "manque précédent : " & arrZMNUMEN0(I).MNUMENPRE
                End If
    
            Else
                arrZMNUMEN0(I).Hierarchie = arrZMNUMEN0(K).Hierarchie & Format(I, "00000")
                arrZMNUMEN0(I).Niveau = arrZMNUMEN0(K).Niveau + 1
            End If
        End If
    Next I
    
Loop Until blnOk

If kHierarchie >= 5 Then MsgBox "voir en fin de liste", vbInformation, "Options orphelines"

ReDim Preserve arrZMNUMEN0(wNb + 1)
ReDim Preserve arrZMNUOPT0(wNb + 1)

Exit Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, App_Debug

End Sub

'---------------------------------------------------------
Public Function rsZMNUMEN0_GetBuffer(rsAdo As ADODB.Recordset, rsZMNUMEN0 As typeZMNUMEN0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZMNUMEN0_GetBuffer = Null

rsZMNUMEN0.MNUMENETB = rsAdo("MNUMENETB")
rsZMNUMEN0.MNUMENREF = rsAdo("MNUMENREF")
rsZMNUMEN0.MNUMENGRP = rsAdo("MNUMENGRP")
rsZMNUMEN0.MNUMENPRE = rsAdo("MNUMENPRE")
rsZMNUMEN0.MNUMENORD = rsAdo("MNUMENORD")
rsZMNUMEN0.MNUMENCOD = rsAdo("MNUMENCOD")
rsZMNUMEN0.MNUMENOIA = rsAdo("MNUMENOIA")
rsZMNUMEN0.MNUMENJOQ = rsAdo("MNUMENJOQ")

rsZMNUMEN0.Method = ""
rsZMNUMEN0.Niveau = 0
rsZMNUMEN0.Hierarchie = ""

Exit Function

Error_Handler:

rsZMNUMEN0_GetBuffer = Error

End Function

