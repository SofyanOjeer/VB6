Attribute VB_Name = "sqlZSWIRAL0"
Option Explicit

Public Sub incrementeSAV()
Dim strChaine As String
Dim fic As Long
Dim Nb As Long

    'désactivé le 24/10/2018 suite à migration vers SAAPROD-VM
    Exit Sub
    fic = FreeFile
    Open paramSAA_Data_Archive & "\SAV_ZSWIRAL0\nb.txt" For Input As #fic
    Do Until EOF(fic)
        Line Input #fic, strChaine
        If InStr(Trim(strChaine), "nbzswiral0=") > -1 Then
            Nb = Mid(Trim(strChaine), 12)
            Exit Do
        End If
    Loop
    Close #fic
    Nb = Nb + 1
    fic = FreeFile
    Open paramSAA_Data_Archive & "\SAV_ZSWIRAL0\zswiral0_" & CStr(Nb) & ".txt" For Output As #fic
    Close #fic
    fic = FreeFile
    Open paramSAA_Data_Archive & "\SAV_ZSWIRAL0\nb.txt" For Output As #fic
    Print #fic, "nbzswiral0=" & CStr(Nb)
    Close #fic

End Sub

Private Sub insertSAV(zxsql As String)
Dim strChaine As String
Dim fic As Long
Dim Nb As Long

    'désactivé le 24/10/2018 suite à migration vers SAAPROD-VM
Exit Sub
On Error Resume Next

    fic = FreeFile
    Open paramSAA_Data_Archive & "\SAV_ZSWIRAL0\nb.txt" For Input As #fic
    Do Until EOF(fic)
        Line Input #fic, strChaine
        If InStr(Trim(strChaine), "nbzswiral0=") > -1 Then
            Nb = Mid(Trim(strChaine), 12)
            Exit Do
        End If
    Loop
    Close #fic
    fic = FreeFile
    Open paramSAA_Data_Archive & "\SAV_ZSWIRAL0\zswiral0_" & CStr(Nb) & ".txt" For Append As #fic
    Print #fic, zxsql
    Close #fic
    
End Sub
Public Function sqlZSWIRAL0_Insert(newY As typeZSWIRAL0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlZSWIRAL0_Insert = Null

xSet = " ("
xValues = " values("

' Détecter les modifications
'===================================================================================

xSet = xSet & ",SWIRALDON": xValues = xValues & " ,'" & Text_Apostrophe(Trim(newY.SWIRALDON)) & "'"
xSet = xSet & ",SWIRALETA": xValues = xValues & " ," & newY.SWIRALETA
xSet = xSet & ",SWIRALMES": xValues = xValues & " ,'" & newY.SWIRALMES & "'"

Mid$(xSet, 3, 1) = " "
Mid$(xValues, 10, 1) = " "

xSql = "Insert into " & paramIBM_Library_SAB & ".ZSWIRAL0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsSab_Update = cnSab_Update.Execute(xSql, Nb)
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlZSWIRAL0_Insert = "Erreur màj : " & Error
    Exit Function
End If

Call insertSAV(xSql)
Call FEU_VERT
 
Exit Function
Error_Handler:
    sqlZSWIRAL0_Insert = Error
End Function



Public Sub ZSWIRAL0_Est_Vide()
Dim rsDenis As ADODB.Recordset
Dim destinataires As String
Dim lMessage As String
Dim xSql As String
Dim Nb As Long

    On Error Resume Next
    xSql = "select count(SWIRALDON) from " & paramIBM_Library_SAB & ".ZSWIRAL0"
    Set rsDenis = cnsab.Execute(xSql, Nb)
    If Not rsDenis.EOF Then
        Nb = CLng(rsDenis(0).value)
    End If
    rsDenis.Close
    Set rsDenis = Nothing
    If Nb > 0 Then
        destinataires = "foucart.m@bia-paris.fr;ligot.p@bia-paris.fr;rosillette.d@bia-paris.fr"
        lMessage = vbCrLf & "ZSWIRAL0 n'est pas vide => " & CStr(Nb) & " enregistrement(s) !"
        lMessage = lMessage & vbCrLf & arrZSWIRAL0(1).SWIRALDON
        Call Email_Standard(destinataires, "@SAA_ENTRANT", lMessage, True, "")
    End If
    
End Sub


