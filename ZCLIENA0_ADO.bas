Attribute VB_Name = "adoZCLIENA0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCLIENA0_PutBuffer(rsAdo As ADODB.Recordset, rszCLIENA0 As typeZCLIENA0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCLIENA0_PutBuffer = Null

rsAdo("CLIENAETB") = rszCLIENA0.CLIENAETB
rsAdo("CLIENACLI") = rszCLIENA0.CLIENACLI
rsAdo("CLIENAAGE") = rszCLIENA0.CLIENAAGE
rsAdo("CLIENAETA") = rszCLIENA0.CLIENAETA
rsAdo("CLIENARA1") = rszCLIENA0.CLIENARA1
rsAdo("CLIENARA2") = rszCLIENA0.CLIENARA2
rsAdo("CLIENASIG") = rszCLIENA0.CLIENASIG
rsAdo("CLIENASRN") = rszCLIENA0.CLIENASRN
rsAdo("CLIENASRT") = rszCLIENA0.CLIENASRT
rsAdo("CLIENADNA") = rszCLIENA0.CLIENADNA
rsAdo("CLIENAREG") = rszCLIENA0.CLIENAREG
rsAdo("CLIENANAT") = rszCLIENA0.CLIENANAT
rsAdo("CLIENARSD") = rszCLIENA0.CLIENARSD
rsAdo("CLIENARES") = rszCLIENA0.CLIENARES
rsAdo("CLIENAECO") = rszCLIENA0.CLIENAECO
rsAdo("CLIENAACT") = rszCLIENA0.CLIENAACT
rsAdo("CLIENAPAI") = rszCLIENA0.CLIENAPAI
rsAdo("CLIENACRD") = rszCLIENA0.CLIENACRD
rsAdo("CLIENAADM") = rszCLIENA0.CLIENAADM
rsAdo("CLIENAATR") = rszCLIENA0.CLIENAATR
rsAdo("CLIENABIL") = rszCLIENA0.CLIENABIL
rsAdo("CLIENACAT") = rszCLIENA0.CLIENACAT
rsAdo("CLIENACOT") = rszCLIENA0.CLIENACOT
rsAdo("CLIENACHQ") = rszCLIENA0.CLIENACHQ
rsAdo("CLIENADAT") = rszCLIENA0.CLIENADAT
rsAdo("CLIENASAC") = rszCLIENA0.CLIENASAC
rsAdo("CLIENAGEO") = rszCLIENA0.CLIENAGEO
rsAdo("CLIENAENT") = rszCLIENA0.CLIENAENT
rsAdo("CLIENAMES") = rszCLIENA0.CLIENAMES
rsAdo("CLIENAPAY") = rszCLIENA0.CLIENAPAY
rsAdo("CLIENAFIL") = rszCLIENA0.CLIENAFIL
rsAdo("CLIENABIM") = rszCLIENA0.CLIENABIM
rsAdo("CLIENADOU") = rszCLIENA0.CLIENADOU
rsAdo("CLIENALI1") = rszCLIENA0.CLIENALI1
rsAdo("CLIENALI2") = rszCLIENA0.CLIENALI2
rsAdo("CLIENAEXT") = rszCLIENA0.CLIENAEXT
rsAdo("CLIENACOL") = rszCLIENA0.CLIENACOL
rsAdo("CLIENATIE") = rszCLIENA0.CLIENATIE
rsAdo("CLIENASEL") = rszCLIENA0.CLIENASEL
rsAdo("CLIENAPCS") = rszCLIENA0.CLIENAPCS
rsAdo("CLIENACRE") = rszCLIENA0.CLIENACRE
Exit Function

Error_Handler:

rsZCLIENA0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZCLIENA0_AddNew(rsAdo As ADODB.Recordset, rszCLIENA0 As typeZCLIENA0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZCLIENA0_AddNew = Null
rsAdo.AddNew
adoZCLIENA0_AddNew = rsZCLIENA0_PutBuffer(rsAdo, rszCLIENA0)
rsAdo.Update

Exit Function

Error_Handler:

adoZCLIENA0_AddNew = Error

End Function
