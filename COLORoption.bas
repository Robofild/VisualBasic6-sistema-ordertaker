Attribute VB_Name = "Personalize"
Option Explicit


Public Function colorOPT() As String
Form444.CommonDialog1.ShowColor
colorOPT = Form444.CommonDialog1.color

End Function


