Attribute VB_Name = "srvYSSISAM0"
Option Explicit

Type typeYSSISAM0
    
    SSISAMETA       As Integer
    SSISAMREF       As Long
    SSISAMGRP       As String
    SSISAMCLA       As String
    SSISAMAPP       As String
    SSISAMCOD       As String
    SSISAMAGE       As Integer
    SSISAMSER       As String
    SSISAMSSE       As String
    SSISAMOPE       As String
    SSISAMNAT       As String
    SSISAMPRD       As String
    SSISAMAUT       As String
    SSISAMFON       As String
    SSISAMDON       As String
    SSISAMCAI       As String
    SSISAMMON       As Currency
    SSISAMDEV       As String
    SSISAMDLY       As String
    SSISAMPRO       As String
    SSISAMCLI       As String
    SSISAMEIC       As String
    SSISAMSDD       As String
    SSISAMDRO       As String
    SSISAMSUC       As String
    SSISAMPRC       As String
    SSISAMNJ1       As String
    SSISAMTJ1       As String
    SSISAMNJ2       As String
    SSISAMTJ2       As String
    SSISAMPRA       As String
    SSISAMECH       As String
    SSISAMUIDX      As String
    SSISAMUIDD      As String
    SSISAMTLNK      As Long
    SSISAMYFCT      As String
    SSISAMYUSR      As String
    SSISAMYAMJ      As Long
    SSISAMYHMS      As Long
    SSISAMYVER      As Long
End Type

'---------------------------------------------------------
Public Function rsYSSISAM0_GetBuffer(rsAdo As ADODB.Recordset, rsYSSISAM0 As typeYSSISAM0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYSSISAM0_GetBuffer = Null

rsYSSISAM0.SSISAMETA = rsAdo("SSISAMETA")
rsYSSISAM0.SSISAMREF = rsAdo("SSISAMREF")
rsYSSISAM0.SSISAMGRP = rsAdo("SSISAMGRP")
rsYSSISAM0.SSISAMCLA = rsAdo("SSISAMCLA")
rsYSSISAM0.SSISAMAPP = rsAdo("SSISAMAPP")
rsYSSISAM0.SSISAMCOD = rsAdo("SSISAMCOD")
rsYSSISAM0.SSISAMAGE = rsAdo("SSISAMAGE")
rsYSSISAM0.SSISAMSER = rsAdo("SSISAMSER")
rsYSSISAM0.SSISAMSSE = rsAdo("SSISAMSSE")
rsYSSISAM0.SSISAMOPE = rsAdo("SSISAMOPE")
rsYSSISAM0.SSISAMNAT = rsAdo("SSISAMNAT")
rsYSSISAM0.SSISAMPRD = rsAdo("SSISAMPRD")
rsYSSISAM0.SSISAMAUT = rsAdo("SSISAMAUT")
rsYSSISAM0.SSISAMFON = rsAdo("SSISAMFON")
rsYSSISAM0.SSISAMDON = rsAdo("SSISAMDON")
rsYSSISAM0.SSISAMCAI = rsAdo("SSISAMCAI")
rsYSSISAM0.SSISAMMON = rsAdo("SSISAMMON")
rsYSSISAM0.SSISAMDEV = rsAdo("SSISAMDEV")
rsYSSISAM0.SSISAMDLY = rsAdo("SSISAMDLY")
rsYSSISAM0.SSISAMPRO = rsAdo("SSISAMPRO")
rsYSSISAM0.SSISAMCLI = rsAdo("SSISAMCLI")
rsYSSISAM0.SSISAMEIC = rsAdo("SSISAMEIC")
rsYSSISAM0.SSISAMSDD = rsAdo("SSISAMSDD")
rsYSSISAM0.SSISAMDRO = rsAdo("SSISAMDRO")
rsYSSISAM0.SSISAMSUC = rsAdo("SSISAMSUC")
rsYSSISAM0.SSISAMPRC = rsAdo("SSISAMPRC")
rsYSSISAM0.SSISAMNJ1 = rsAdo("SSISAMNJ1")
rsYSSISAM0.SSISAMTJ1 = rsAdo("SSISAMTJ1")
rsYSSISAM0.SSISAMNJ2 = rsAdo("SSISAMNJ2")
rsYSSISAM0.SSISAMTJ2 = rsAdo("SSISAMTJ2")
rsYSSISAM0.SSISAMPRA = rsAdo("SSISAMPRA")
rsYSSISAM0.SSISAMECH = rsAdo("SSISAMECH")
rsYSSISAM0.SSISAMUIDX = rsAdo("SSISAMUIDX")
rsYSSISAM0.SSISAMUIDD = rsAdo("SSISAMUIDD")
rsYSSISAM0.SSISAMTLNK = rsAdo("SSISAMTLNK")
rsYSSISAM0.SSISAMYFCT = rsAdo("SSISAMYFCT")
rsYSSISAM0.SSISAMYUSR = rsAdo("SSISAMYUSR")
rsYSSISAM0.SSISAMYAMJ = rsAdo("SSISAMYAMJ")
rsYSSISAM0.SSISAMYHMS = rsAdo("SSISAMYHMS")
rsYSSISAM0.SSISAMYVER = rsAdo("SSISAMYVER")

Exit Function

Error_Handler:

rsYSSISAM0_GetBuffer = Error

End Function

