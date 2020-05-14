Attribute VB_Name = "maskCep"
Option Explicit


Public Function CepMask(cepNum As String) As String
CepMask = Format(cepNum, "00\.000\-000")

End Function
