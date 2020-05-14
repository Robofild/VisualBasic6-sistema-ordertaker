Attribute VB_Name = "ValidarCGCCPF"
Option Explicit


Public Function CpfFormat(digitacaoCpf As String)
'CalculaCGC = Format(Numero, "000\.000\.000\-00")

CpfFormat = Format(digitacaoCpf, "000\.000\.000\-00")
End Function

Public Function CnpjFormat(digitacaoCnpj As String)

CnpjFormat = Format(digitacaoCnpj, "00\.000\.000\/0000\-00")
End Function

Public Function CalculaCGC(numero As String) As String

Dim i As Integer
Dim prod As Integer
Dim mult As Integer
Dim digito As Integer

If Not IsNumeric(numero) Then
   CalculaCGC = ""
   Exit Function
End If

mult = 2
For i = Len(numero) To 1 Step -1
  prod = prod + Val(Mid(numero, i, 1)) * mult
  mult = IIf(mult = 9, 2, mult + 1)
Next

digito = 11 - Int(prod Mod 11)
digito = IIf(digito = 10 Or digito = 11, 0, digito)

CalculaCGC = Trim(Str(digito))

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
Dim i As Integer 'utilizada nos FOR... NEXT
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
For i = 2 To 10
    strCaracter = Right(strcampo, i - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * i
    lngSoma = lngSoma + intMais
Next i
dblDivisao = lngSoma / 11

lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig1 = 0
Else
    intDig1 = 11 - intResto
End If

strcampo = strcampo & intDig1 'concatena o CPF com o primeiro digito verificador
lngSoma = 0
intNumero = 0
intMais = 0
'Inicia cálculos do 2º dígito
For i = 2 To 11
    strCaracter = Right(strcampo, i - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * i
    lngSoma = lngSoma + intMais
Next i
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


