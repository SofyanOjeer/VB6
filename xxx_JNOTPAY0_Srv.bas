Attribute VB_Name = "srvJNOTPAY0"
'---------------------------------------------------------
Option Explicit
Type typeJNOTPAY0
    JORCV                   As Long
    JOSEQN                  As Long
    JRNBIATRN               As Long

    NOTPAYISO   As String * 2  ' code ISO pays
    NOTPAYSEQ   As Long        ' N° séquence (info)
    
    NOTPAYPROV  As String * 1  ' Provisionable = 'P'
    NOTPAYCOFA  As String * 2  ' notation coface
    NOTPAYCOFK  As String * 1  ' notation coface Auto / Manuel
    NOTPAYCOFD  As Long        ' DATE maj
    NOTPAYOCDE  As String * 1  ' notation OCDE
    NOTPAYOCDK  As String * 1  ' notation OCDE Auto / Manuel
    NOTPAYOCDD  As Long        ' DATE maj
    NOTPAYSP    As String * 4  ' notation S & P
    NOTPAYSPK   As String * 1  ' notation S & P Auto / Manuel
    NOTPAYSPD  As Long        ' DATE maj
    NOTPAYCEG   As Long        ' critère événement grave
    NOTPAYBIAN  As String * 3  ' notation BIA
    NOTPAYBIAK  As String * 1  ' notation BIA Auto / Manuel
    NOTPAYBIAD  As Long        ' DATE maj
    NOTPAYTAUX  As Double      ' taux BIA
    NOTPAYFISC  As String * 2  ' taux fisc
    NOTPAYTXT   As String * 32 ' commentaire
    NOTPAYXAMJ  As Long        ' DATE maj
    NOTPAYXHMS  As Long        ' heure maj
    NOTPAYXUSR  As String * 10 ' utilisateur maj
    
End Type

Public Function sqlJNOTPAY0_Read(oldY As typeJNOTPAY0)
Dim X As String, xSql As String, Nb As Long
Dim V

On Error GoTo Error_Handler
sqlJNOTPAY0_Read = Null

xSql = "select * from " & paramIBM_Library_SABSPE & ".JNOTPAY0 " _
       & " where JORCV = '" & oldY.JORCV & "'" _
       & " and   JOSEQN = " & oldY.JOSEQN

Set rsSab = cnsab.Execute(xSql)

If rsSab.EOF Then
    sqlJNOTPAY0_Read = "? inconnu"
Else
    V = rsJNOTPAY0_GetBuffer(rsSab, oldY)
    If Not IsNull(V) Then sqlJNOTPAY0_Read = "? srvJNOTPAY0_GetBuffer"
End If
 
Exit Function
Error_Handler:
    sqlJNOTPAY0_Read = Error
End Function


'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsJNOTPAY0_GetBuffer(rsADO As ADODB.Recordset, rsJNOTPAY0 As typeJNOTPAY0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsJNOTPAY0_GetBuffer = Null

rsJNOTPAY0.JORCV = rsADO("JORCV")
rsJNOTPAY0.JOSEQN = rsADO("JOSEQN")
rsJNOTPAY0.JRNBIATRN = rsADO("JRNBIATRN")

rsJNOTPAY0.NOTPAYISO = rsADO("NOTPAYISO")
rsJNOTPAY0.NOTPAYSEQ = rsADO("NOTPAYSEQ")
rsJNOTPAY0.NOTPAYPROV = rsADO("NOTPAYPROV")
rsJNOTPAY0.NOTPAYCOFA = rsADO("NOTPAYCOFA")
rsJNOTPAY0.NOTPAYCOFK = rsADO("NOTPAYCOFK")
rsJNOTPAY0.NOTPAYCOFD = rsADO("NOTPAYCOFD")
rsJNOTPAY0.NOTPAYOCDE = rsADO("NOTPAYOCDE")
rsJNOTPAY0.NOTPAYOCDK = rsADO("NOTPAYOCDK")
rsJNOTPAY0.NOTPAYOCDD = rsADO("NOTPAYOCDD")
rsJNOTPAY0.NOTPAYSP = rsADO("NOTPAYSP")
rsJNOTPAY0.NOTPAYSPK = rsADO("NOTPAYSPK")
rsJNOTPAY0.NOTPAYSPD = rsADO("NOTPAYSPD")
rsJNOTPAY0.NOTPAYCEG = rsADO("NOTPAYCEG")
rsJNOTPAY0.NOTPAYBIAN = rsADO("NOTPAYBIAN")
rsJNOTPAY0.NOTPAYBIAK = rsADO("NOTPAYBIAK")
rsJNOTPAY0.NOTPAYBIAD = rsADO("NOTPAYBIAD")
rsJNOTPAY0.NOTPAYTAUX = rsADO("NOTPAYTAUX")
rsJNOTPAY0.NOTPAYFISC = rsADO("NOTPAYFISC")
rsJNOTPAY0.NOTPAYTXT = rsADO("NOTPAYTXT")
rsJNOTPAY0.NOTPAYXAMJ = rsADO("NOTPAYXAMJ")
rsJNOTPAY0.NOTPAYXHMS = rsADO("NOTPAYXHMS")
rsJNOTPAY0.NOTPAYXUSR = rsADO("NOTPAYXUSR")

Exit Function

Error_Handler:

rsJNOTPAY0_GetBuffer = Error

End Function




