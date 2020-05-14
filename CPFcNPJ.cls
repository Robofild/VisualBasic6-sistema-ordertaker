VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPFcNPJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim OPERACAO As Boolean



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


'---------------------------------------------------------

Public Function CalculaCGC(Numero As String) As String

Dim I As Integer
Dim prod As Integer
Dim mult As Integer
Dim digito As Integer
Dim identifique As Integer

If Not IsNumeric(Numero) Then
   CalculaCGC = ""
   Exit Function
End If

mult = 2
For I = Len(Numero) To 1 Step -1
  prod = prod + Val(Mid(Numero, I, 1)) * mult
  mult = IIf(mult = 9, 2, mult + 1)
Next

digito = 11 - Int(prod Mod 11)
digito = IIf(digito = 10 Or digito = 11, 0, digito)

CalculaCGC = Trim(Str(digito))

If CalculaCGC = 0 Then


 If Len(Numero) > 11 Then
    MsgBox "cnpj valido"
 OPERACAO = 101
 Else
 
 MsgBox "cPF valido"
 OPERACAO = 104

 End If
 
 Else
 
 If Len(Numero) > 11 Then
 MsgBox "cnpj nao valido"
 Else
 MsgBox "Cpf  Nao valido"
 End If

End If


End Function
Public Function ValidaCGC(CGC As String) As Boolean
If CalculaCGC(Left(CGC, 12)) <> Mid(CGC, 13, 1) Then
   ValidaCGC = False
   Exit Function
End If

If CalculaCGC(Left(CGC, 13)) <> Mid(CGC, 14, 1) Then
   ValidaCGC = False
   Exit Function
End If

ValidaCGC = True

End Function
'2- Validar CPF
Function calculacpf(CPF As String) As Boolean
'Esta rotina foi adaptada da revista Fórum Access
On Error GoTo Err_CPF
Dim I As Integer 'utilizada nos FOR... NEXT
Dim strcampo As String 'armazena do CPF que será utilizada para o cálculo
Dim strCaracter As String 'armazena os digitos do CPF da direita para a esquerda
Dim intNumero As Integer 'armazena o digito separado para cálculo (uma a um)
Dim intMais As Integer 'armazena o digito específico multiplicado pela sua base
Dim lngSoma As Long 'armazena a soma dos digitos multiplicados pela sua base(intmais)
Dim dblDivisao As Double 'armazena a divisão dos digitos*base por 11
Dim lngInteiro As Long 'armazena inteiro da divisão
Dim intResto As Integer 'armazena o resto
Dim intDig1 As Integer 'armazena o 1º digito verificador
Dim intDig2 As Integer 'armazena o 2º digito verificador
Dim strConf As String 'armazena o digito verificador

lngSoma = 0
intNumero = 0
intMais = 0
strcampo = Left(CPF, 9)

'Inicia cálculos do 1º dígito
For I = 2 To 10
    strCaracter = Right(strcampo, I - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * I
    lngSoma = lngSoma + intMais
Next I
dblDivisao = lngSoma / 11

lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig1 = 0
   
   
Else
    intDig1 = 11 - intResto
  MsgBox "invalido!"
     
End If

strcampo = strcampo & intDig1 'concatena o CPF com o primeiro digito verificador
lngSoma = 0
intNumero = 0
intMais = 0
'Inicia cálculos do 2º dígito
For I = 2 To 11
    strCaracter = Right(strcampo, I - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * I
    lngSoma = lngSoma + intMais
Next I
dblDivisao = lngSoma / 11
lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig2 = 0
Else
    intDig2 = 11 - intResto
End If
strConf = intDig1 & intDig2
'Caso o CPF esteja errado dispara a mensagem
If strConf <> Right(CPF, 2) Then
    calculacpf = False
Else
    calculacpf = True
End If
Exit Function

Exit_CPF:
    Exit Function
Err_CPF:
    MsgBox Error$
    Resume Exit_CPF
End Function


Public Function CpfFormat(digitacaoCpf As String)
'CalculaCGC = Format(Numero, "000\.000\.000\-00")

CpfFormat = Format(digitacaoCpf, "000\.000\.000\-00")
End Function

Public Function CnpjFormat(digitacaoCnpj As String)

CnpjFormat = Format(digitacaoCnpj, "00\.000\.000\/0000\-00")
End Function




