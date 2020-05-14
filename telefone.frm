VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Form3 
   Caption         =   "Informe o telefone"
   ClientHeight    =   3195
   ClientLeft      =   9465
   ClientTop       =   5430
   ClientWidth     =   3975
   Icon            =   "telefone.frx":0000
   LinkTopic       =   "Form3"
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3975
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Avançar"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      _Version        =   393216
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####-####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   600
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Enter"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Telefone:"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim counter  As Integer
Dim numero As String
Dim chaveMataprocesso As Boolean





Private Sub Command1_Click()
Proceguir
End Sub

Private Sub Command2_Click()

'MaskEdBox1.Text = nor
  'txtTexto.SelStart = 0
   'txtTexto.SelLength = Len(txtTexto.Text)
End Sub

Private Sub Form_Load()
Dialog.Hide
Form2.Show
Form2.Visible = False
Form23.Show
Form23.Visible = False
Form500.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
matarprocesso
End Sub

Private Sub MaskEdBox1_Change()
chaveMataprocesso = True

           
               











End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
If chaveMataprocesso = True Then
 matarprocesso
 chaveMataprocesso = False
 End If
 End If
 If KeyAscii = 13 Then
Proceguir
End If

End Sub

Private Sub MaskEdBox1_LostFocus()

 matarprocesso

End Sub

Private Sub Text1_Change()
  Form2.Text4.Text = Text1.Text
  Form2.Text11 = "inicio"
    Form2.Text12 = "inicio"
            'Form2.Command2.Enabled = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'TodoCaptarTelefone
Dim contador As Integer

If Text1.Text <> "" Then
' counter = Left(Text1.Text, 1)
contador = Len(Text1.Text)
 
                                    Text1.Text = Format(Text1.Text, "####-####")
                                    If contador = 8 Then
                                    Form2.Show
                                    Form2.Text4.Text = Text1.Text
                        '            If Form2.Text1.Enabled = True Then
                                    Form2.Text1.SetFocus
                                     End If
            'If KeyAscii = 8 And contador <= 5 Then
           '  If KeyAscii = 8 Or KeyAscii = 0 Then
            '           Text1.Text = ""
                       
                      ' Unload Form2
                       'Unload (Form2)
                       
                      ' matarprocesso
                     ' Text1.SetFocus
              Else
                
                     '   If counter = 9 Then
                      '       If Text1.SelStart = 5 Then Text1.SelText = "-"
                             ' Text1.MaxLength = 10

                                    
                                   ' Form2.Command2.Enabled = True
                             'End If
                             
                      End If
                      'If counter = 2 Or counter = 3 Then
                       '     If Text1.SelStart = 4 Then Text1.SelText = "-"
                        '     Text1.MaxLength = 9
                         '    If contador = 8 Then
                          '   Form2.Show
                           '  Form2.Text4.Text = Text1.Text
                            ' If Form2.Text1.Enabled = True Then
                             'Form2.Text1.SetFocus
                             'End If
                             'End If
                       'End If
                        'If counter <> 9 And counter <> 3 And counter <> 2 Then
                         '   MsgBox "Seu numero pode estar errado !", , "Digito 9?"
                       'End If
                     
             'End If

'End If
                        
  'Form2.Text4.Text = Text1.Text

End Sub

Public Sub matarprocesso()
Dim appName As String

Dim Comando As String
appName = "Atendimento.exe"
Comando = "TASKKILL -F -IM " & appName
Shell Comando
appName = "Atendimento.exe"
Comando = "TASKKILL -F -IM " & appName
Shell Comando
End Sub


Public Sub TodoCaptarTelefone()
'Dim contador As Integer
'
'If Text1.Text <> "" Then
' counter = Left(Text1.Text, 1)
'contador = Len(Text1.Text)
'
'            'If KeyAscii = 8 And contador <= 5 Then
'             If KeyAscii = 8 Or KeyAscii = 0 Then
'                       Text1.Text = ""
'
'                       Unload Form2
'                       'Unload (Form2)
'
'                       matarprocesso
'                      Text1.SetFocus
'              Else
'
'                        If counter = 9 Then
'                             If Text1.SelStart = 5 Then Text1.SelText = "-"
'                              Text1.MaxLength = 10
'                             If contador = 9 Then
'                                    Form2.Show
'                                    Form2.Text4.Text = Text1.Text
'                                    If Form2.Text1.Enabled = True Then
'                                    Form2.Text1.SetFocus
'                                    End If
'
'                                   ' Form2.Command2.Enabled = True
'                             End If
'
'                      End If
'                      If counter = 2 Or counter = 3 Then
'                            If Text1.SelStart = 4 Then Text1.SelText = "-"
'                             Text1.MaxLength = 9
'                             If contador = 8 Then
'                             Form2.Show
'                             Form2.Text4.Text = Text1.Text
'                             If Form2.Text1.Enabled = True Then
'                             Form2.Text1.SetFocus
'                             End If
'                             End If
'                       End If
'                        If counter <> 9 And counter <> 3 And counter <> 2 Then
'                            MsgBox "Seu numero pode estar errado !", , "Digito 9?"
'                       End If
'
'             End If
'
'End If
'
'  'Form2.Text4.Text = Text1.Text
End Sub

Public Sub Proceguir()
Unload Form22
If MaskEdBox1.Text <> "____-____" Then
                       
                       Unload Form2
                       'Unload (Form2)
                       
                       matarprocesso
                      MaskEdBox1.SetFocus
     
                                    Form2.Visible = True
                                    Form2.Text4.Text = MaskEdBox1.Text
                                    If Form2.Text1.Enabled = True Then
                                    Form2.Text1.SetFocus
                                    End If
                                    
                                    Form2.Visible = True
                         
End If
                        
  'Form2.Text4.Text = Text1.Text
End Sub
