VERSION 5.00
Begin VB.Form Form100 
   Caption         =   "Form9"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   LinkTopic       =   "Form9"
   ScaleHeight     =   4815
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "f1"
      Height          =   735
      Left            =   4920
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "SIMULADOR DE TECLAGEM"
      Height          =   855
      Left            =   1080
      TabIndex        =   7
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "case Web"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "sql"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "criar mask em tempo de execulsao"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "sleep"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ANIMAÇÃO"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "COMSERVER"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LINHA DO BANCO"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub Command1_Click()


'Adodc3.RecordSource = ""

'Adodc3.CommandType = adCmdText

'Adodc3.RecordSource = "SELECT * FROM `Cardapio` WHERE `idCardapio` = ' " & indice & " '"

'Adodc3.Refresh
'Text4.Text = Format(Adodc3.Recordset("valor").Value, "Currency")
 
'Text4.Text = Format(Adodc3.Recordset("valor").Value, "Currency")
 
'End Sub

'Private Sub Command2_Click()

'ConServer
'
'Dim sql As String
'Dim rs As New ADODB.Recordset
'Set rs = New ADODB.Recordset
'Set rs.ActiveConnection = con

'sql = "SELECT * FROM `Cardapio_Tipo`,`TituloCardapio` WHERE `tipos` LIKE '" & nomeCategoria & "' AND `nometitulo` LIKE '" & nomeTitulo & "'"
'rs.Open sql



'If rs.BOF = False Then
'IdTituloVar = rs.Fields("idTituloCardapio").Value
'iDCategoriaVar = rs.Fields("idCardapio_Tipo").Value

'End If
' Set rs = Nothing
'End Sub

'Private Sub Command3_Click()
'Private Sub Command8_Click()

 '   On Error GoTo erro
  '  Dim con As Object
    'Set con = CreateObject("ADODB.Connection")
   ' Animation1.Open "C:\Program Files\visual studio 6\COMMON\GRAPHICS\AVIS\BLUR8.AVI"
    'Animation1.AutoPlay = True
    'DoEvents
    'con.Open "Provider=SQLOLEDB;" & _
     '        "Data Source=ServidorQueNaoExiste;" & _
      '       "User Id=sa;" & _
       '      "Password=123456;" & _
        '     "Initial Catalog=BancoQueNaoExiste;"
  '  Animation1.Close
   ' con.Close
    'Set con = Nothing
    'MsgBox "Conexão efetuada com sucesso!", vbInformation, "Programação On-Line"
    'Exit Sub
'erro:
 '   MsgBox Err.Description, vbExclamation, "Programação On-Line"
  '  Animation1.Close

Private Sub Command2_Click()
ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

rs.CursorLocation = adUseClient
  sql = "SELECT COUNT(DISTINCT `Titulo`) FROM `Cardapio` WHERE 1"
  rs.Open sql


 Set rs = Nothing
End Sub

'End Sub
'End Sub
Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
'Private Declare Sub 'Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
 
'usage
''Sleep 3000
'causes program to pause for 3 seconds
End Sub

Private Sub Command5_Click()

'Private Sub Text1_KeyPress(KeyAscii As Integer)
'Dim contador As Integer
'If Text1.Text <> "" Then
'counter = Left(Text1.Text, 1)
'contador = Len(Text1.Text)
 
    'Select Case KeyAscii
        'Case 48 To 57
            'If Text1.SelStart = 0 Then Text1.SelText = "("
            'If Text1.SelStart = 3 Then Text1.SelText = ") "
 '       If KeyAscii = 8 And contador <= 5 Then
  '          Text1.Text = ""
    '        Text1.SetFocus
   '     Else
     '       If counter = 9 Then
      '            If Text1.SelStart = 5 Then Text1.SelText = "-"
       '     End If
         '    If counter = 2 Or counter = 3 Then
        '          If Text1.SelStart = 4 Then Text1.SelText = "-"
          '  End If
         'End If
       ' End Select
'End If
End Sub


Private Sub Command6_Click()
'        Adodc1.Recordset("CEP").Value = tXTcEP.Text
'Adodc1.Recordset.Update
'
'ConServer
'
'Dim sql As String
'Dim rs As New ADODB.Recordset
'Set rs = New ADODB.Recordset
'Set rs.ActiveConnection = con
''entrada de dados
'    If telefoneTipo = 9 Then
'    sql = " INSERT INTO `robofi61_order_taker`.`Cli_celular` (`telefone_celular`, `fk_cliente`) VALUES (' " & Text4.Text & " ', 'NULL')"
'    rs.Open sql
'
'    sql = "INSERT INTO `clientes` (`idclientes`, `tele_fixo`, `fk_nome`, `fk_endereco`, `fk_celular`, `fk_pedidos`) VALUES (NULL, NULL, NULL, NULL, NULL, NULL)"
'    rs.Open sql
'    Else
'    sql = "INSERT INTO `Cli_telefixo` ( `telefone_fixo`, `fk_cliente`) VALUES ( ' " & Text4.Text & " ', NULL)"
'     rs.Open sql
'     sql = "INSERT INTO `clientes` (`idclientes`, `tele_fixo`, `fk_nome`, `fk_endereco`, `fk_celular`, `fk_pedidos`) VALUES (NULL, ' " & Text4.Text & " ', NULL, NULL, NULL, NULL)"
'     rs.Open sql
'
'    End If
'    ' apt
'    If Text3.Text <> "" Then
'    sql = "INSERT INTO `cli_apt` (`id`, `numero_apt`, `fk_bloco`) VALUES (NULL, ' " & Text3.Text & " ', NULL)"
'     rs.Open sql
'    End If
'    'bloco
'    If Text5.Text <> "" Then
'    sql = "INSERT INTO `cli_bloco` (`id`, `bloco`, `fk_apt`, `fk_endereco`) VALUES (NULL, ' " & Text5.Text & " ', NULL, NULL)"
'     rs.Open sql
'    End If
'
'  'cep
'   If tXTcEP.Text <> "" Then
'    sql = "INSERT INTO `cli_cepIndex` (`id`, `fk_numeros`, `cli_cepIndexcol`) VALUES (NULL,  ' " & tXTcEP.Text & " ', NULL)"
'    rs.Open sql
'   Else
'   sql = "INSERT INTO `cli_CepNegado` (`id`, `Endereco`, `Bairro`, `uf`, `Cidade`) VALUES (NULL, ' " & Text6.Text & " ', ' " & Text8.Text & " ' , ' " & Text9.Text & " ', ' " & Text10.Text & " ')"
'   rs.Open sql
'
'   End If
'   If Text1.Text <> "" Then
'   sql = "INSERT INTO `cli_nomes` (`id_nomes`, `nome`, `fk_cliente`, `fk_endereco`, `fk_celulares`, `id_fixo`, `creatd`, `id_user`) VALUES (NULL, ' " & Text1.Text & " ', NULL, NULL, NULL, NULL, NULL, NULL)"
'    rs.Open sql
'   End If
'
'     If txtnumero.Text <> "" Then
'     sql = "INSERT INTO `Cli_endereco` (`id`, `numero`, `id_cep`, `fk_apts`, `fk_blocos`, `outros`) VALUES (NULL, ' " & txtnumero.Text & " ', ' " & tXTcEP.Text & " ', NULL, NULL, ' " & Text2.Text & " ')"
'     rs.Open sql
'     End If
'
'             ' reconect completando
'
'             Dim fkcliente, fkclientenome, FKtelefonesfixos, fkapt, fkbloco, fkcep, fknome, fknumero, fkcelular As Integer
'
'             'codigo nome cliente
'             If Text1.Text <> "" Then
'             sql = "SELECT `id_nomes` FROM `cli_nomes` ORDER BY `cli_nomes`.`id_nomes` DESC LIMIT 1"
'             rs.Open sql
'             fkclientenome = rs.Fields("id_nomes").Value
'             rs.Close
'
'             'retonao o code da tabela cliente refernciada
'             sql = "SELECT `idclientes` FROM `clientes` ORDER BY `idclientes` DESC LIMIT 1"
'             rs.Open sql
'             fkcliente = rs.Fields("idclientes").Value
'             rs.Close
'
'
'
'            End If
'            'telefones fixos do cliente
'
'             If Text4.Text <> "" And telefoneTipo <> 9 Then
'             sql = "SELECT `id` FROM `Cli_telefixo` WHERE `telefone_fixo` ORDER BY `id` DESC LIMIT 1"
'            rs.Open sql
'             FKtelefonesfixos = rs.Fields("id").Value
'             rs.Close
'             Else
'             sql = "SELECT `id` FROM `Cli_celular` WHERE `id` ORDER BY `id`DESC LIMIT 1"
'             rs.Open sql
'             fkcelular = rs.Fields("id").Value
'             rs.Close
'             End If
'
'
'            ' apt
'            If Text3.Text <> "" Then
'             sql = "SELECT `id` FROM `cli_apt` WHERE `id` ORDER BY `id` DESC LIMIT 1"
'             rs.Open sql
'             fkapt = rs.Fields("id").Value
'             rs.Close
'            End If
'
'            'bloco
'            If Text5.Text <> "" Then
'             sql = "SELECT `id` FROM `cli_bloco` WHERE `id` ORDER BY `id` DESC LIMIT 1"
'             rs.Open sql
'             fkbloco = rs.Fields("id").Value
'             rs.Close
'            End If
'
'            'retorna fk cep
'           If tXTcEP.Text <> "" Then
'            sql = "SELECT `id` FROM `cli_cepIndex` WHERE `id` ORDER BY `id`DESC LIMIT 1"
'              rs.Open sql
'             fkcep = rs.Fields("id").Value
'             rs.Close
'             Else
'             sql = "SELECT `id` FROM `cli_CepNegado` WHERE `id` ORDER BY `id`DESC LIMIT 1"
'              rs.Open sql
'             fkcep = rs.Fields("id").Value
'             rs.Close
'           End If
'           'retirba fk nomees
'           If Text1.Text <> "" Then
'             sql = "SELECT `id_nomes` FROM `cli_nomes` WHERE `id_nomes`ORDER by `id_nomes`DESC LIMIT 1"
'             rs.Open sql
'             fknome = rs.Fields("id_nomes").Value
'             rs.Close
'           End If
'
'             If txtnumero.Text <> "" Then
'              sql = "SELECT `id` FROM `Cli_endereco` WHERE `id` ORDER by `id`DESC LIMIT 1"
'             rs.Open sql
'             fknumero = rs.Fields("id").Value
'            rs.Close
'
'             End If
'
'
'
'    'Dim fkcliente, FKtelefonesfixos, fkapt, fkbloco, fkcep, fknome, fknumero As Integer
'            'completa a tabela de clientes
'
'            sql = "UPDATE `clientes` SET `fk_nome`=' " & fkclientenome & " ',`fk_endereco`=' " & fknumero & " ',`fk_celular`=' " & fkcelular & " ',`fk_pedidos`='0' WHERE `idclientes` ORDER BY `idclientes`DESC LIMIT 1"
'             rs.Open sql
'             'completa endereco
'             If tXTcEP.Text = "" Then
'             sql = "UPDATE `Cli_endereco` SET `id_cep`=' " & fkcep & " ',`fk_apts`=' " & fkapt & " ',`fk_blocos`=' " & fkbloco & " '  WHERE `id` ORDER BY `id`DESC LIMIT 1"
'              rs.Open sql
'              Else
'              sql = "UPDATE `Cli_endereco` SET `fk_apts`=' " & fkapt & " ',`fk_blocos`=' " & fkbloco & " '  WHERE `id` ORDER BY `id`DESC LIMIT 1"
'              rs.Open sql
'              End If
'
'              'cep
'              sql = "UPDATE `cli_cepIndex` SET `fk_numeros`=' " & fknumero & " ' WHERE `id` ORDER BY `id`DESC LIMIT 1"
'              rs.Open sql
'              'celular
'              sql = "UPDATE `Cli_celular` SET `fk_cliente`=' " & fkcliente & " 'WHERE `id` ORDER BY `id`DESC LIMIT 1"
'              rs.Open sql
'               'bloco
'              sql = "UPDATE `cli_bloco` SET `fk_apt`=' " & fkapt & " ',`fk_endereco`=' " & fknumero & " ' WHERE  `id` ORDER BY `id`DESC LIMIT 1"
'              rs.Open sql
'               'apt
'              sql = "UPDATE `cli_apt` SET `fk_bloco`=' " & fkbloco & " ' WHERE `id` ORDER BY `id` DESC LIMIT 1"
'              rs.Open sql
'                 'nomes
'              sql = "UPDATE `cli_nomes` SET `fk_cliente`=' " & fkcliente & " ',`fk_endereco`=' " & fknumero & " ' ,`fk_celulares`=' " & fkcelular & " ',`id_fixo`=' " & FKtelefonesfixos & " ' ,`creatd`=' " & Now() & " ',`id_user`='0' WHERE `id_nomes` ORDER BY `id_nomes`DESC LIMIT 1"
'              rs.Open sql
'                 'telefixo
'              sql = "UPDATE `Cli_telefixo` SET `fk_cliente`=' " & fkcliente & " ' WHERE `id` ORDER BY `id`DESC LIMIT 1"
'              rs.Open sql
'
'
' Set rs = Nothing
'
'MsgBox "Novo Cliente cadastrado com sucesso!", , "Novo Cadastro"
'

End Sub

Private Sub Command7_Click()
'On Error Resume Next
'
'        Select Case Index
'            Case 0 'botão Imprimir
'
'            Case 1 'Botão Visualizar
'
'            Case 2 'botão Configurar
'
'            Case 3 'botão propriedades
'
'        End Select
End Sub

Private Sub Command8_Click()
'https://docs.microsoft.com/pt-br/office/vba/language/reference/user-interface-help/sendkeys-statement
Dim ReturnValue, i
ReturnValue = Shell("CALC.EXE", 1)    ' Run Calculator.
'AppActivate ReturnValue     ' Activate the Calculator.
'For I = 1 To 100    ' Set up counting loop.
 '   SendKeys I & "{+}", True    ' Send keystrokes to Calculator
'Next I    ' to add each value of I.
'SendKeys "=", True    ' Get grand total.
SendKeys "%{F4}", True    ' Send ALT+F4 to close Calculator.
End Sub

Private Sub Command9_Click()
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
'  If KeyCode = vbKeyF1 Then
'    MsgBox "F1"
'  End If

End Sub

