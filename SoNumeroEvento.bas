Attribute VB_Name = "SoNumeroEvento"
Option Explicit


   Function SoNumeroseventos(ByVal KeyAscii As Integer) As Integer

     If InStr("1234567890", Chr(KeyAscii)) = 0 Then

        SoNumeros = 0

     Else

        SoNumeros = KeyAscii

      End If

 

     Select Case KeyAscii

        Case 8

        SoNumeros = KeyAscii

        Case 13

        SoNumeros = KeyAscii

        Case 32
        'nao permitir espaços em branco
        'SoNumeros = KeyAscii
       
        End Select

   End Function

