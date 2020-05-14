VERSION 5.00
Begin VB.Form Form405 
   BackColor       =   &H8000000A&
   Caption         =   "Form9"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12195
   Icon            =   "Form405.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   5145
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      Text            =   "Text8"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   7440
      TabIndex        =   12
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Text            =   "Text7"
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   8280
      TabIndex        =   10
      Text            =   "Text6"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   8280
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   4440
      Width           =   6135
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Form405.frx":058A
      Left            =   600
      List            =   "Form405.frx":05A9
      TabIndex        =   4
      Text            =   "Forma de pagamento"
      Top             =   240
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Text            =   "75,10"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Concluir o pagamento"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   2880
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "Troco"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Valor"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Dinheiro"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Form405"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim controleTrasito As Boolean

Private Sub Combo1_Click()

Dim Index As Integer
          
            If Combo1.Text <> "Forma de pagamento" Then
            If Combo1.ListIndex = 0 Then
            Text1.Visible = True
            Text3.Visible = True
            Label1.Visible = True
            Label3.Visible = True
            
            
            Else
            Text3.Visible = False
            Text1.Visible = False
            Label1.Visible = False
            Label3.Visible = False
            
            End If
            
            Index = Combo1.ListIndex
            
              Select Case Index
                        Case 0 '(Entregar) Dinheiro
                         If Text3.Text <> "" And Text3.Text <> "R$ 0,00" Then
                          Text4.Text = "TROCO DE  => " & Text3.Text
                          Else
                          Text4.Text = "SEM TROCO"
                          End If
                          Text8.Text = "0"
                           
                          
                        Case 1 '(Entregar) PG Cartão
                         'informçao adicional
                         Text4.Text = "LEVAR MAQUINA DE CARTAO"
                         'chave de index tomada de decisão
                         Text8.Text = "1"
                         'controleTrasito = True
                         formatUniversal
                        
                        Case 2 '(Entregar) Ticket de alimentação
                           'informçao adicional
                         Text4.Text = "PAGAMENTO EM TICKETs"
                         'chave de index tomada de decisão
                         Text8.Text = "2"
                         'controleTrasito = True
                         formatUniversal

                         
                        Case 3 '(Entregar) Pago!
                          'informçao adicional
                         Text4.Text = "PAGO SO ENTREGAR"
                         'chave de index tomada de decisão
                         Text8.Text = "3"
                         'controleTrasito = True
                         formatUniversal
                        
                        
                        Case 4 '(Entregar) Pagará Depois
                             'informçao adicional
                         Text4.Text = "SO ENTREGAR -ANOTAR"
                         'chave de index tomada de decisão
                         Text8.Text = "4"
                         'controleTrasito = True
                         formatUniversal
                         
                         
                         Case 5 '(Balcão) CLIENTE ESTA ESPERANDO
                         'informçao adicional
                         Text4.Text = "CLIENTE NA LOJA ESPERANDO"
                         'chave de index tomada de decisão
                         Text8.Text = "5"
                         'controleTrasito = True
                         formatUniversal
                        
                        
                        Case 6 '(Balcão) Pago! vem busca
                         'informçao adicional
                         Text4.Text = "PAGO CLIENTE VEM BUSCAR"
                         'chave de index tomada de decisão
                         Text8.Text = "6"
                         'controleTrasito = True
                         formatUniversal
                        
                        
                        Case 7 '(Balcão) Pagar na hora que buscar
                         'informçao adicional
                         Text4.Text = "PAGARA NA HORA QUE BUSCAR"
                         'chave de index tomada de decisão
                         Text8.Text = "7"
                         'controleTrasito = True
                         formatUniversal
                        
                        
                        Case 8 '(Balcão) Pagará Depois
                         'informçao adicional
                         Text4.Text = "SO ENTREGAR - ANOTAR "
                         'chave de index tomada de decisão
                         Text8.Text = "8"
                         'controleTrasito = True
                         formatUniversal
                            
                           
                           
                           
                           
                  End Select
 Command2.Enabled = True
Else
 Command2.Enabled = False
End If
End Sub

Private Sub Command1_Click()
If Text1.Text = "" And Text5 <> "" Then
            'DBLcurre (Text2.Text)
            Text1.Text = Text2.Text
            Text5.Text = Text2.Text
            Text6.Text = Text2.Text
            caucular
            End If
valor1 = Replace(Text5, ".", ",")
valor2 = Replace(Text6, ".", ",")
valor3 = Format(valor1 - valor2, "currency")
Text3.Text = Format(valor3, "currency")


Text1.Text = Format(Text1.Text, "currency")
Text2.Text = Format(Text2.Text, "currency")
Text3.Text = Format(Text3.Text, "currency")
  Call Combo1_Click

End Sub

Private Sub Command2_Click()
Form404.Show
Form404.Text3.Text = Text5.Text
Form404.Text4.Text = Text6.Text
Form404.Text5.Text = Text7.Text
Form404.Text6.Text = Text8.Text
Form404.Text7.Text = Text4.Text
Form404.Command2.Enabled = True
'Call Form404.Command6_Click
Form405.Hide
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
KeyAscii = 0
Text1.Text = ""
End If
If KeyAscii = 13 Then
KeyAscii = 0
            If Text1 <> "" Then
            Text1.Text = Replace(Text1.Text, ".", ",")
            Text1.Text = Format(Text1.Text, "currency")
            
             
            Text5.Text = Text1
            Text6.Text = Text2
            caucular
            Else
            
            Text5.Text = Text2.Text
            Text6.Text = Text2.Text
            caucular
            
            
            End If
End If
End Sub

Private Sub Text1_LostFocus()
If Text1 <> "" Then
Text1.Text = Replace(Text1.Text, ".", ",")
Text1.Text = Format(Text1.Text, "currency")

 
Text5.Text = Text1
Text6.Text = Text2
caucular
Else

Text5.Text = Text2.Text
Text6.Text = Text2.Text
caucular


End If
'valor3 = valor2 - valor1
'Text7.Text = valor3



End Sub

Private Sub Text2_Change()
Text2.Text = Format(Text2.Text, "currency")
End Sub

Private Sub Text2_LostFocus()
Text2.Text = Format(Text2.Text, "currency")
End Sub

Private Sub Text3_Change()
Text7.Text = Text3.Text
End Sub

Private Sub Text5_Change()
Text5 = DBLcurre(Text5.Text)
'Text5 = CDbl(Text5.Text)
End Sub

Private Sub Text6_Change()
Text6 = DBLcurre(Text6.Text)
'ext6 = CDbl(Text6.Text)
End Sub

Private Sub Text7_Change()
Text7 = DBLcurre(Text7.Text)

End Sub

Public Sub caucular()
If Text1.Text = "" And Text5 <> "" Then
            'DBLcurre (Text2.Text)
            Text1.Text = Text2.Text
            Text5.Text = Text2.Text
            Text6.Text = Text2.Text
            caucular
            End If
valor1 = Replace(Text5, ".", ",")
valor2 = Replace(Text6, ".", ",")
valor3 = Format(valor1 - valor2, "currency")
Text3.Text = Format(valor3, "currency")


Text1.Text = Format(Text1.Text, "currency")
Text2.Text = Format(Text2.Text, "currency")
Text3.Text = Format(Text3.Text, "currency")
  Call Combo1_Click


'
'
'valor1 = Replace(Text5, ".", ",")
'valor2 = Replace(Text6, ".", ",")
'valor3 = Format(valor1 - valor2, "currency")
'Text3.Text = Format(valor3, "currency")
'Call Combo1_Click
End Sub

Public Sub formatUniversal()


Text5.Text = Text2.Text
Text6.Text = Text2.Text
Text7.Text = 0
'caucular


'End If
   
End Sub

Private Sub Text8_Change()
If Text1.Visible = True And Text1.Text <> "" Then
Text5.Text = Text1.Text
Text6.Text = Text2.Text
Text7.Text = Text3.Text
End If
End Sub
