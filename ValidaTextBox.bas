Attribute VB_Name = "ValidaTextBox"
Option Explicit


 

  Function SoLETRAS(ByVal KeyAscii As Integer) As Integer

        'Transformar letras minusculas em Maiúsculas

     KeyAscii = Asc(UCase(Chr(KeyAscii)))

       ' Intercepta um código ASCII recebido e admite somente letras

          If InStr("AÃÁBCÇDEÉÊFGHIÍJKLMNOPQRSTUÚVWXYZ", Chr(KeyAscii)) = 0 Then
    
             SoLETRAS = 0
    
         Else
    
             SoLETRAS = KeyAscii
    
          End If

 

   ' teclas adicionais permitidas

    If KeyAscii = 8 Then SoLETRAS = KeyAscii ' Backspace

    If KeyAscii = 13 Then SoLETRAS = KeyAscii ' Enter

    If KeyAscii = 32 Then SoLETRAS = KeyAscii ' Espace

End Function

 

 

   Function SoNumeros(ByVal KeyAscii As Integer) As Integer

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

 
Function SoNumerosvirgulaPonto(ByVal KeyAscii As Integer) As Integer


     If InStr("1234567890.,", Chr(KeyAscii)) = 0 Then

        SoNumerosvirgulaPonto = 0

     Else

        SoNumerosvirgulaPonto = KeyAscii

      End If

 

     Select Case KeyAscii

        Case 8

        SoNumerosvirgulaPonto = KeyAscii

        Case 13

        SoNumerosvirgulaPonto = KeyAscii

        Case 32
        'nao permitir espaços em branco
        'SoNumeros = KeyAscii
        
        End Select

   End Function
   
   Function SoPonto(ByVal KeyAscii As Integer) As Integer


     If InStr("1234567890.", Chr(KeyAscii)) = 0 Then

        SoPonto = 0

     Else

        SoPonto = KeyAscii

      End If

 

     Select Case KeyAscii

        Case 8

        SoPonto = KeyAscii

        Case 13

        SoPonto = KeyAscii

        Case 32
        'nao permitir espaços em branco
        'SoNumeros = KeyAscii
        
        End Select

   End Function

