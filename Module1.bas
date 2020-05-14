Attribute VB_Name = "Module1"
Option Explicit

Dim valotransf As String

Public Function DBLcurre(valorEmString) As String
valotransf = Replace(valorEmString, ",", ".")
valotransf = Replace(valotransf, "R$", "")
DBLcurre = Trim(valotransf)

End Function
