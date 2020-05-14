Attribute VB_Name = "masktelefone"
Option Explicit




Public Function phoneformat(digitacaoPhone As String)


Dim tipodemask, constroimaskparaTelefone As Integer
tipodemask = Len(digitacaoPhone)
Select Case tipodemask
            Case 11
          
             constroimaskparaTelefone = 11
            Case 9
             
             constroimaskparaTelefone = 9
            Case 8
            constroimaskparaTelefone = 8
             
            Case Else
            constroimaskparaTelefone = 0
              
    End Select





Select Case constroimaskparaTelefone
            Case 11
            digitacaoPhone = Format(digitacaoPhone, "(##)#####-####")
           
              
            Case 9
             digitacaoPhone = Format(digitacaoPhone, "(31)#####-####")
            Case 8
            'constroimaskparaTelefone = 8
              digitacaoPhone = Format(digitacaoPhone, "(31)####-####")
            Case Else
            'constroimaskparaTelefone = 0
            digitacaoPhone = Format(digitacaoPhone, "###-####")
    End Select
    
    
   phoneformat = digitacaoPhone
End Function




