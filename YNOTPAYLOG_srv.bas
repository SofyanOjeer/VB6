Attribute VB_Name = "srvYNOTPAYLOG"
'---------------------------------------------------------
Option Explicit
Type typeYNOTPAYLOG

    NOTPAYLOGD  As Long        ' DATE maj
    NOTPAYLOGH  As Long        ' heure maj
    NOTPAYLOGU  As String * 10 ' utilisateur maj
    NOTPAYLOGS   As Long        ' N° séquence (info)
    NOTPAYLOGK   As String * 10 ' code action
    NOTPAYLOGX   As String * 64 ' commentaire
    
    
End Type
Public Function sqlYNOTPAYLOG_Insert(newY As typeYNOTPAYLOG)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYNOTPAYLOG_Insert = Null

xSet = " (NOTPAYLOGS"
xValues = " values(" & newY.NOTPAYLOGS

' Insertion :
'===================================================================================
If newY.NOTPAYLOGD <> 0 Then xSet = xSet & ",NOTPAYLOGD": xValues = xValues & ", " & newY.NOTPAYLOGD
If newY.NOTPAYLOGH <> 0 Then xSet = xSet & ",NOTPAYLOGH": xValues = xValues & ", " & newY.NOTPAYLOGH

'===================================================================================

If Trim(newY.NOTPAYLOGK) <> "" Then xSet = xSet & ",NOTPAYLOGK": xValues = xValues & ", '" & Replace(Trim(newY.NOTPAYLOGK), "'", "''") & "'"
If Trim(newY.NOTPAYLOGX) <> "" Then xSet = xSet & ",NOTPAYLOGX": xValues = xValues & ", '" & Replace(Trim(newY.NOTPAYLOGX), "'", "''") & "'"
If newY.NOTPAYLOGU <> "" Then xSet = xSet & ",NOTPAYLOGU": xValues = xValues & ", '" & Replace(newY.NOTPAYLOGU, "'", "''") & "'"
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YNOTPAYLOG" & xSet & ")" & xValues & ")"

Set rsSab = cnsab.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYNOTPAYLOG_Insert = "Erreur màj : " & newY.NOTPAYLOGX & " " & newY.NOTPAYLOGS
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYNOTPAYLOG_Insert = Error
End Function
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYNOTPAYLOG_GetBuffer(rsAdo As ADODB.Recordset, rsYNOTPAYLOG As typeYNOTPAYLOG)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYNOTPAYLOG_GetBuffer = Null

rsYNOTPAYLOG.NOTPAYLOGD = rsAdo("NOTPAYLOGD")
rsYNOTPAYLOG.NOTPAYLOGH = rsAdo("NOTPAYLOGH")
rsYNOTPAYLOG.NOTPAYLOGU = rsAdo("NOTPAYLOGU")
rsYNOTPAYLOG.NOTPAYLOGS = rsAdo("NOTPAYLOGS")
rsYNOTPAYLOG.NOTPAYLOGK = rsAdo("NOTPAYLOGK")
rsYNOTPAYLOG.NOTPAYLOGX = rsAdo("NOTPAYLOGX")

Exit Function

Error_Handler:

rsYNOTPAYLOG_GetBuffer = Error

End Function








