VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form404 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Verifique com atenção o pedido"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   Icon            =   "visualizacaodopedidos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   6990
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   5760
      TabIndex        =   47
      Text            =   "?"
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Height          =   375
      Left            =   6360
      Picture         =   "visualizacaodopedidos.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   8040
      TabIndex        =   43
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pagar"
      Height          =   615
      Left            =   120
      Picture         =   "visualizacaodopedidos.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   8520
      TabIndex        =   37
      Text            =   "0"
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   11520
      TabIndex        =   36
      Text            =   "Text8"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   1095
      Left            =   7920
      TabIndex        =   35
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   11160
      TabIndex        =   34
      Text            =   "Text7"
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   11160
      TabIndex        =   33
      Text            =   "Text6"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   11160
      TabIndex        =   32
      Text            =   "Text5"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   11160
      TabIndex        =   31
      Text            =   "Text4"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   11160
      TabIndex        =   30
      Text            =   "Text3"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "visualizacaodopedidos.frx":1E06
      Height          =   495
      Left            =   1080
      TabIndex        =   28
      Top             =   7560
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cancelar "
      Height          =   615
      Left            =   120
      Picture         =   "visualizacaodopedidos.frx":1E1B
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6240
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8640
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=robofi61_order_taker"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "robofi61_order_taker"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "at_itens"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      Picture         =   "visualizacaodopedidos.frx":281D
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Adicionar"
      Height          =   615
      Left            =   120
      Picture         =   "visualizacaodopedidos.frx":321F
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remover"
      Height          =   615
      Left            =   120
      Picture         =   "visualizacaodopedidos.frx":3D21
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1080
      TabIndex        =   12
      Top             =   8280
      Width           =   5775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   9000
      Width           =   5775
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "visualizacaodopedidos.frx":4723
      Height          =   4335
      Left            =   1080
      TabIndex        =   0
      Top             =   2520
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7646
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "QUANTIDADE / DESCRIÇÃO"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8160
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=robofi61_order_taker"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "robofi61_order_taker"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "at_frete"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   5640
      TabIndex        =   46
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Edição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label26 
      Caption         =   "Ediçao"
      Height          =   255
      Left            =   8040
      TabIndex        =   42
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label25 
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label24 
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label23 
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label22 
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label21 
      Caption         =   "FRETE..>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   29
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label Label20 
      Caption         =   "Label20"
      Height          =   255
      Left            =   5040
      TabIndex        =   27
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
      Height          =   375
      Left            =   1320
      TabIndex        =   26
      Top             =   9600
      Width           =   2535
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      Height          =   255
      Left            =   2640
      TabIndex        =   23
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "LOJA ENTREGA :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   22
      Top             =   2160
      Width           =   1530
   End
   Begin VB.Label Label15 
      Caption         =   "Label15"
      Height          =   255
      Left            =   5280
      TabIndex        =   21
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ENDEREÇO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   5
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   375
      Left            =   2040
      TabIndex        =   18
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "TELEFONE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Top             =   840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   0
      Picture         =   "visualizacaodopedidos.frx":4738
      Stretch         =   -1  'True
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "OBSERV...>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   8520
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "TOTAL  ->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   9000
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "DATA /HORA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   9600
      Width           =   1185
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "OPERADOR :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3960
      TabIndex        =   7
      Top             =   9600
      Width           =   1170
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "REFERNCIA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   6
      Top             =   1680
      Width           =   1125
   End
   Begin VB.Label Label4 
      Caption         =   "CLIENTE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Nº"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "PEDIDO :                                                        "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   6600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "NOME DA SUA EMPRESA"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form404"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cancelar As Boolean
Dim numerodopedido As Integer
Dim FreteouIten As Integer
Dim dataSelecionado As Integer
Dim Revalidar As Boolean
Dim idProdutO As Integer
Dim chaveSOBREescrevaoCupon As Boolean

Private Sub Command1_Click()
Form401.Show
Form401.Text16.Text = 0
Form401.Text16.Visible = False
Form401.Text31.Text = 1
Revalidar = True
'repassar o numero do pedido para edição
If Text10.Text = "1" Then
Form401.Label15.Visible = True
Form401.Label15.Caption = Label18.Caption
Form401.Text29.Text = " 1"
Form401.Text31.Text = Text10.Text
End If
Form401.Show
Form404.Hide
Command2.Enabled = False

End Sub

Private Sub Command2_Click()
Dim numPedido As Integer
Dim NpTitulo As String
numPedido = Label18.Caption
If Revalidar = True Then
'Printer.Print " -----------------------------------------------"
'Printer.Print "MODIFICACAO DO PEDIDO:  ", numPedido
'Printer.Print ")"
'Printer.Print ""
'Printer.Print ""
'Printer.Print " -----------------------------------------------"
'Printer.Print "PEDIDO MODIFICADO AGUARDE NOVA COMANDA"
'Printer.Print " -----------------------------------------------"
'Printer.Print "MODIFICACAO  DESTE PEDIDO ", numPedido
'Printer.Print " -----------------------------------------------"
'Printer.Print " "
Revalidar = False
ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient



sql = "UPDATE `at_contadorDePedidos` SET `situacaoImpressao` = '1' WHERE `at_contadorDePedidos`.`id` =  '" & numPedido & "'"
                            rs.Open sql
 Set rs = Nothing
End If
Call Command6_Click
'imprimirCupon
Cancelar = True

Unload Form404
Unload Form401
Unload Form402
Form3.Show
End Sub
Private Sub cmdImprimir()
PrintLabels2 (11)
'Dim Printer As New Printer


Printer.EndDoc

End Sub
Private Sub PrintLabels2(iAmount As Integer)
Dim i As Integer
Dim X As Printer
 
For Each X In Printers
 
    If X.DeviceName = "\\TELEAMARELINHO\Generic / Text Only" Then
    ' Set printer as system default.
    Set Printer = X
    ' Stop looking for a printer.
    Exit For
    End If
 
Next
 
For i = 1 To iAmount
    With Printer
        .Height = 1440
        .Width = 1440 * 3
        .CurrentY = 100
        .CurrentX = 700
        .FontSize = 8
        .FontBold = True
        .Font = "Courier New"
    End With
'    Printer.Print UCase(Inventory_Recs(1).Inventory_Style_Number & " " & Trim$(cboSize.Text) & " " & txtDescription.Text)
'    Printer.CurrentY = 280
'    Printer.CurrentX = 700
'    Printer.Print UCase(txtExtendedDesc.Text)
'    Printer.CurrentY = 460
'    Printer.CurrentX = 700
'    Printer.Font = "WASP 39 H"
'    Printer.FontBold = False
'    Printer.FontSize = 44
   ' stBarcode = Replace(Mid$(stBarcode, 1, 10), " ", "_")
    'Printer.Print "*" & stBarcode & "*"
    Printer.EndDoc
Next i
 
End Sub

Private Sub Command3_Click()
If chaveSOBREescrevaoCupon = True Then
metSobrescrevaoCUpom
End If

Form407.Text8 = Text1
If Text10.Text = 1 Then
Form406.Text1 = 5
Form406.Show

End If
If Form406.Text1 = 1 Then
Form406.Text1 = 0
Else
If Form406.Text1 <> 5 Then
Form406.Text1 = 1
End If
End If

Form500.Hide
Form406.Caption = "Pagamento do pedido  Nº" & Form404.Label18.Caption
Form407.Text1.Text = Text1.Text
Form407.Label1.Caption = Form404.Label18.Caption
End Sub

Private Sub Command4_Click()
RemoverFreteouIten (dataSelecionado)

Atualize


End Sub

Private Sub Command5_Click()
Unload Form500
Dim resp As Integer
'Form2.Hide
'Form405.Hide
'Form403.Hide
If Text9.Text = "0" Then
Cancelar = False
Else
Cancelar = True

End If
If Cancelar = False Then
        resp = MsgBox("Você não  consolidou o  pedido falta resolver o pagameto , se você decidir fechar; o pedido será cancelado ! Deseja continuar mesmo assim?", vbYesNo, "Cancelar o pedido ?")
        If resp = 6 Then
        CancelarPedido
        Form401.Text26.Text = 0
        Unload Form2
        Form3.Show
        Else
        Form401.Text26.Text = 1
        End If
Else
Unload Form402
'Unload Form405
Unload Form403
Unload Form403

Form2.Show
'Unload Me

End If


End Sub

Private Sub Command6_Click()
ConServer
Dim operador As String
Dim valorFRete As String
Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

rs.CursorLocation = adUseClient
'CommonDialog1.CancelError = True
'trarar erro
On Error GoTo error



valorFRete = DBLcurre(DataGrid2.Columns(0).Value)

  
  sql = "INSERT INTO `at_Cupon` (`id`, `nomeEmpresa`, `numPedido`, `endereco`, `telefone`, `referencia`, `loja`, `fk_itens`, `valor_frete`, `obsvacoes`, `total`, `datahora`, `operador`, `valorRecebido`, `valrorPago`, `troco`, `observacoes2`, `formadepagamento`)" & _
  "VALUES (NULL, '" & Label1.Caption & "', '" & Label18.Caption & "','" & Label13.Caption & "', '" & Label15.Caption & "', '" & Label14.Caption & "', '" & Label17.Caption & "', '" & Label18.Caption & "','" & valorFRete & "' , '" & Text2.Text & "', '" & Text4.Text & "','" & Label19.Caption & "', '" & Label20.Caption & "', '" & Text3.Text & "', '" & Text4.Text & "', '" & Text5.Text & "', '" & Text7.Text & "', '" & Text6.Text & "')"
  rs.Open sql
  
 

    
    
   
   
   
   
   
   
   
   
   
   
   
   
 Set rs = Nothing


Exit Sub

error:
valorFRete = 0

  
  sql = "INSERT INTO `at_Cupon` (`id`, `nomeEmpresa`, `numPedido`, `endereco`, `telefone`, `referencia`, `loja`, `fk_itens`, `valor_frete`, `obsvacoes`, `total`, `datahora`, `operador`, `valorRecebido`, `valrorPago`, `troco`, `observacoes2`, `formadepagamento`)" & _
  "VALUES (NULL, '" & Label1.Caption & "', '" & Label18.Caption & "','" & Label13.Caption & "', '" & Label15.Caption & "', '" & Label14.Caption & "', '" & Label17.Caption & "', '" & Label18.Caption & "','" & valorFRete & "' , '" & Text2.Text & "', '" & Text4.Text & "','" & Label19.Caption & "', '" & Label20.Caption & "', '" & Text3.Text & "', '" & Text4.Text & "', '" & Text5.Text & "', '" & Text7.Text & "', '" & Text6.Text & "')"
  rs.Open sql

Exit Sub




End Sub

Private Sub Command7_Click()
Atualize
End Sub

Private Sub DataGrid1_Click()
dataSelecionado = 1
idProdutO = DataGrid1.Columns(3).Value
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
    Call Command3_Click
  End If
    If KeyCode = vbKeyF2 Then
   Call Command1_Click
  End If
    If KeyCode = vbKeyF3 Then
    Call Command4_Click
  End If
    If KeyCode = vbKeyF4 Then
    Call Command5_Click
  End If

End Sub

Private Sub DataGrid2_Change()
contabilize
End Sub

Private Sub DataGrid2_Click()
dataSelecionado = 2

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF1 Then
    Call Command3_Click
  End If
    If KeyCode = vbKeyF2 Then
   Call Command1_Click
  End If
    If KeyCode = vbKeyF3 Then
    Call Command4_Click
  End If
    If KeyCode = vbKeyF4 Then
    Call Command5_Click
  End If

End Sub

Private Sub Form_Load()
Form407.Show
Form407.Visible = False
Form404.Label20.Caption = Form1.StatusBar1.Panels(2).Text
If Form2.Text7.Text = "true" Then
Carregarvaloresdeeditor
End If
verificarsepossuiaFrete
Unload Form500
Form3.Hide
Dialog.Hide
Label12.Caption = Form2.Text1.Text
If funqualEndereco = True Then
repaceEnderecoAutomatico
Else
repaceEnderecoManual
End If
'revogada
'Label15.Caption = Form2.Text4.Text
'Label14.Caption = Form2.Text2.Text
'Label17.Caption = Form2.DataCombo1.Text

Form404.DataGrid1.Columns(2).Width = 1800
Atualize
End Sub

Public Sub repaceEnderecoManual()
If Form2.Text3.Text <> "" Then
Label13.Caption = Form2.Text6.Text & ", nº " & Form2.txtnumero & " - " & Form2.Text3.Text & ", " & Form2.Text8.Text & " - " & Form2.Text10.Text & ""
ElseIf Form2.Text5 <> "" Then
Label13.Caption = Form2.Text6.Text & ", nº " & Form2.txtnumero & " / " & Form2.Text5.Text & ", " & Form2.Text8.Text & " - " & Form2.Text10.Text & ""
ElseIf Form2.Text5 <> "" And Form2.Text3 <> "" Then
Label13.Caption = Form2.Text6.Text & ", nº " & Form2.txtnumero & " / " & Form2.Text5.Text & "AP " & Form2.Text3.Text & ", " & Form2.Text8.Text & " - " & Form2.Text10.Text & ""
Else
Label13.Caption = Form2.Text6.Text & ", nº " & Form2.txtnumero & ", " & Form2.Text8.Text & " - " & Form2.Text10.Text & ""
End If
End Sub



Public Sub repaceEnderecoAutomatico()
If Form2.Text3 <> "" Then
Label13.Caption = Form2.txtRua.Text & ", nº " & Form2.txtnumero & "-  " & Form2.Text3.Text & ", " & Form2.tXTbAIRRO.Text & " - " & Form2.TXTcIDADE.Text & ""
ElseIf Form2.Text5 <> "" Then
Label13.Caption = Form2.txtRua.Text & ", nº " & Form2.txtnumero & " / " & Form2.Text5.Text & ", " & Form2.tXTbAIRRO.Text & " - " & Form2.TXTcIDADE.Text & ""
ElseIf Form2.Text5 <> "" And Form2.Text3 <> "" Then
Label13.Caption = Form2.txtRua.Text & ", nº " & Form2.txtnumero & " / " & Form2.Text5.Text & "AP " & Form2.Text3.Text & ", " & Form2.tXTbAIRRO.Text & " - " & Form2.TXTcIDADE.Text & ""
Else
Label13.Caption = Form2.txtRua.Text & ", nº " & Form2.txtnumero & ", " & Form2.tXTbAIRRO.Text & " - " & Form2.TXTcIDADE.Text & ""
End If
End Sub



Public Function funqualEndereco() As Boolean
If Form2.Text6.Visible = True Then
funqualEndereco = False
Else
funqualEndereco = True
End If
End Function

Public Sub contabilize()
Dim valordefrete As Double
ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
numerodopedido = Label18.Caption
Dim valorDosItens As Double
rs.CursorLocation = adUseClient
  sql = "SELECT SUM(`valor`) FROM `at_itens` WHERE `fk_pedido` = '" & numerodopedido & "'"
  rs.Open sql
  If Adodc2.Recordset.BOF = False Then
  Text8.Text = Adodc2.Recordset.Fields("frete").Value
  valordefrete = Text8.Text
  Else
  valordefrete = 0
  End If
  'CommonDialog1.CancelError = True
'trarar erro
On Error GoTo error



  valorDosItens = rs.Fields("SUM(`valor`)").Value
  
  
 If valordefrete <> 0 Then
   Form404.Text1.Text = Format(valorDosItens + Text8.Text, "currency")
Else
  Form404.Text1.Text = Format(valorDosItens, "currency")
End If
 Set rs = Nothing
    Exit Sub

error:
valorDosItens = 0

 If valordefrete <> 0 Then
   Form404.Text1.Text = Format(valorDosItens + Text8.Text, "currency")
Else
  Form404.Text1.Text = Format(valorDosItens, "currency")
End If
 Set rs = Nothing
Exit Sub
    
    
   
 Set rs = Nothing
End Sub

Public Sub removerIten()
'removerIten
numerodopedido = Label18.Caption
Form404.Adodc1.RecordSource = ""

Form404.Adodc1.CommandType = adCmdText

Form404.Adodc1.RecordSource = "Select * FROM `at_itens` WHERE `fk_pedido` = '" & numerodopedido & "' AND `id` = '" & idProdutO & "'"
Adodc1.Refresh
If Adodc1.Recordset.BOF = False Then

Adodc1.Recordset.Delete
End If
'Adodc1.Refresh


Form404.Adodc1.RecordSource = ""

Form404.Adodc1.CommandType = adCmdText

Form404.Adodc1.RecordSource = "SELECT `Quantidade`,`descrição`,`valor`,`id` FROM `at_itens` WHERE `fk_pedido` = '" & numerodopedido & "'"

Form404.Adodc1.Refresh
'DataGrid2.Columns(0).Value = 0






contabilize

End Sub

Public Sub removerfrete()
Dim resp
resp = MsgBox("Você esta preste a remover o frete esse procedimento não tem volta para recolocar o frete terá que fazer um novo pedido !   mesmo assim deseja continuar?", vbYesNo, "Excluir o Frete?")
If resp = 6 Then
numerodopedido = Label18.Caption
Form404.Adodc2.RecordSource = ""

Form404.Adodc2.CommandType = adCmdText

Form404.Adodc2.RecordSource = "SELECT * FROM `at_frete` WHERE `fk_numPedido`='" & numerodopedido & "'"

'Adodc2.Refresh

Adodc2.Refresh
If Adodc2.Recordset.BOF = False Then
Adodc2.Recordset.Delete
Else
'DataGrid2.Columns(0).Value = 0
End If


Form404.Adodc2.RecordSource = ""

Form404.Adodc2.CommandType = adCmdText

Form404.Adodc2.RecordSource = "SELECT `frete` FROM `at_frete` WHERE `fk_numPedido` ='" & numerodopedido & "'"

Form404.Adodc2.Refresh
End If

contabilize

End Sub

Public Function RemoverFreteouIten(qual As Integer)
If qual = 1 Then
removerIten
'remover item
Else
'remover frere
removerfrete
End If

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Form500
Dim resp As Integer
'Form2.Hide
'Form405.Hide
'Form403.Hide
If Text9.Text = "0" Then
Cancelar = False
Else
Cancelar = True

End If
If Cancelar = False Then
        resp = MsgBox("Você não  consolidou o  pedido falta resolver o pagameto , se você decidir fechar; o pedido será cancelado ! Deseja continuar mesmo assim?", vbYesNo, "Cancelar o pedido ?")
        If resp = 6 Then
        CancelarPedido
        Form401.Text26.Text = 0
        Unload Form2
        Form3.Show
        Else
            'Form401.Text26.Text = 1
        End If
Else
Unload Form402
'Unload Form405
Unload Form403
Unload Form403

'Form2.Show
'Unload Me

End If



End Sub

Private Sub Label27_Click()
Form2.SetFocus
Form2.Text7.Text = "True"
chaveSOBREescrevaoCupon = True
Unload Me


End Sub

Private Sub Text10_Change()
If Text10.Text = 1 Then
Label27.Visible = True
Form401.Text31.Text = 1
Else
Label27.Visible = False
Form401.Text31.Text = 0
End If
End Sub

Private Sub Text11_Change()
If Text11.Text = "2" Then
Label28.Visible = True
Else
Label28.Visible = True
Label28.Caption = "Buscar"
End If
End Sub

Private Sub Text7_Change()
If Text6 >= 5 Then
contabilize
End If
End Sub

Private Sub Text8_Change()
Text6 = DBLcurre(Text6.Text)
End Sub

Public Sub imprimirCupon()

Dim idPiloto As Integer
Dim ContadorPiloto As Integer
Dim PedidoEntrege As Integer
'obter parra o cupom
Dim nomeEmpresa  As String
Dim nomeDoCliente As String
Dim numeroPedido As Integer
Dim endereco As String
Dim telefone As String
Dim referenciaEntrega As String
Dim loja As String
Dim valorFRete As Double
Dim observacoesDopedido As String
Dim TotalDaCompra As Double
Dim Datahora As String
Dim operador As String
Dim ValorRecebido As Double
Dim ValorPago As Double
Dim Troco As Double
Dim ObservacoesparaEntrega As String
Dim formaPagamento As Integer

'itens
Dim QtditensPorLista As Integer
Dim indexB As Integer
Dim QtdP As Integer
Dim descricaoP As String
Dim PrecoP As Double

'fonte padrao
Dim fontepadrao As Integer
Dim OldFont
indexB = 0

                                          
                                          
                                          
                                          
ConServer


Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

rs.CursorLocation = adUseClient
'trasendocontador de pedidos ultimo com contador 1

  sql = "SELECT * FROM `at_contadorDePedidos` WHERE `contador` = 1 ORDER BY `id` DESC LIMIT 1"
  rs.Open sql
  idPiloto = rs.Fields("id").Value
  rs.Close

          ' trazer cupom combinado ao ultimo contador
            sql = "SELECT * FROM `at_Cupon` WHERE `numPedido` = '" & idPiloto & "'ORDER BY `id` DESC LIMIT 1"
            rs.Open sql
              If rs.BOF = False Then
            'Recebendo dados para o cupom
                nomeEmpresa = rs.Fields("nomeEmpresa").Value
                'nomeDoCliente = rs.Fields("nomeCliente").Value
                numeroPedido = rs.Fields("numPedido").Value
                endereco = rs.Fields("endereco").Value
                telefone = rs.Fields("telefone").Value
                referenciaEntrega = rs.Fields("referencia").Value
                loja = rs.Fields("loja").Value
                valorFRete = Format(rs.Fields("valor_frete").Value, "General Number")
                observacoesDopedido = rs.Fields("obsvacoes").Value
                TotalDaCompra = rs.Fields("total").Value
                Datahora = rs.Fields("datahora").Value
                operador = rs.Fields("operador").Value
                ValorRecebido = rs.Fields("valorRecebido").Value
                ValorPago = rs.Fields("valrorPago").Value
                Troco = rs.Fields("troco").Value
                ObservacoesparaEntrega = rs.Fields("observacoes2").Value
                formaPagamento = rs.Fields("formadepagamento").Value
                Else
                numeroPedido = 1
                endereco = "CANCELAR"
                telefone = "CANCELAR"
                referenciaEntrega = "CANCELAR"
                loja = "CANCELAR"
                valorFRete = 0
                observacoesDopedido = "CANCELAR"
                TotalDaCompra = 0
                Datahora = "CANCELAR"
                operador = "CANCELAR"
                ValorRecebido = 0
                ValorPago = 0
                Troco = 0
                ObservacoesparaEntrega = "CANCELAR"
                formaPagamento = 10
                
                
                End If
                
            
            
            
            rs.Close
                          
                          
                          
                          
                          'obter a quantidade de loopings melhor dizer ober quantidade de itens
                            sql = "SELECT COUNT(*) FROM `at_itens` WHERE `fk_pedido` = '" & idPiloto & "'"
                            rs.Open sql
                            QtditensPorLista = rs.Fields("COUNT(*)").Value
                            rs.Close
                                        
   'itens do ultimo cupom combinidos com o contador
                                        
                                          sql = "SELECT * FROM `at_itens` WHERE `fk_pedido` ='" & idPiloto & "'"
                                          rs.Open sql
                                          If rs.BOF = False Then
                                          nomeDoCliente = rs.Fields("fk_cliente").Value
                                          'For indexB = 20 To 1 Step -3
                                          'sb.Append (indexB.ToString)
                                          'sb.Append (" ")
                                          'Next indexB
                                        
                                           
' inicio da impressão dinheiro entrega
'fontepadrao = Printer.FontSize
' Dim PrinterFont As New StdFont
 
OldFont = Printer.FontName   ' Preserva a fonte original.
Printer.Print Tab(10); Format(nomeEmpresa, ">")
Printer.Print
Printer.Print " -----------------------------------------------"

Printer.FontSize = "15"
If formaPagamento <= 4 Then
Printer.Print " ( ENTREGA )       Nº PEDIDO:  "; numeroPedido
Else
        If formaPagamento = 10 Then
        Printer.Print " *******CANCELAR+++++++!!!!!!!!!       Nº PEDIDO:  "; numeroPedido
        Else
        Printer.Print " ( BALCAO )       Nº PEDIDO:  "; numeroPedido
        End If
End If
Printer.FontSize = "10"
Printer.Print " -----------------------------------------------"
Printer.Print "TELEFONE:  "; telefone, nomeDoCliente
Printer.Print "ENDERECO:   "; Format(endereco, ">")
Printer.Print ""
Printer.Print "REFERENCIA: ", Format(referenciaEntrega, ">")
Printer.Print " -----------------------------------------------"
Printer.Print "LOJA:  "; Format(loja, ">")
Printer.Print " -----------------------------------------------"
Printer.Print "QUANTIDADE / DESCRICAO               / PRECO "
Printer.Print " -----------------------------------------------"
Printer.Print " "
                                          For indexB = 1 To QtditensPorLista
                                          QtdP = rs.Fields("Quantidade").Value
                                          descricaoP = rs.Fields("descrição").Value
                                          PrecoP = rs.Fields("valor").Value
                                          Printer.Print " -----------------------------------------------"
                                           If formaPagamento = 10 Then
        Printer.Print " *******  CANCELAR  +++++++!!!!!!!!!       Nº PEDIDO:  "; numeroPedido
        End If
                                          
                                          Printer.Print "  "; QtdP, descricaoP, Spc(10); Format(PrecoP, "Currency")
                                          Printer.Print
                                          If QtditensPorLista = 1 Then
                                            Printer.Print
                                            Printer.Print
                                            Printer.Print
                                            Printer.Print
                                          End If
                                          
                                          Printer.Print
                                          Printer.Print " -----------------------------------------------"
                                          rs.MoveNext
                                          Next indexB

Printer.Print " OBS  "; observacoesDopedido
Printer.Print " -----------------------------------------------"
Printer.Print Tab(15); " SOMATORIA"
Printer.Print
Printer.Print " -----------------------------------------------"
        Select Case formaPagamento
            Case 0 '(Entregar) Dinheiro
                                          Printer.Print "DINHEIRO.............", ; Format(ValorRecebido, "Currency")
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")
                                          Printer.Print
                                          Printer.Print
                                          Printer.Print
                                          Printer.Print "TROCO...................", ; Format(Troco, "Currency")
 
            Case 1 '(Entregar) PG Cartão
                                          Printer.Print "CARTAO"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")

                                          
            Case 2 '(Entregar) Ticket de alimentação
                                          Printer.Print "TICKET"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")

            Case 3 '(Entregar) Pago!
                                          Printer.Print "PAGO"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")


            
            Case 4 ' (Entregar) Pagará depois-Anotar
                                          Printer.Print "ANOTAR"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")

            Case 5 '(Balcão) Cliente aguarda na Loja
                                          Printer.Print "CLIENTE NA LOJA"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")

            Case 6 ' (Balcão) Pago! vem buscar
                                          Printer.Print "VEM BUSCAR PAGO!"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")
            Case 7 ' (Balcão) Pago! vem buscar
                                          Printer.Print "VEM BUSCAR PAGO!"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")
            Case 8 ' (Balcão) Pagar na hora que buscar
                                          Printer.Print "VEM BUSCAR / RECEBER!"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")
            Case 9 ' (Anotar) Pagará Depois
                                          Printer.Print "ANOTAR"
                                          Printer.Print
                                          Printer.Print "TAXA DE ENTREGA -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "TOTAL...................", ; Format(TotalDaCompra, "Currency")
            Case 10 ' (Cancelar) Cancelar
                                          Printer.Print "*******CANCELAR "
                                          Printer.Print
                                          Printer.Print "NÃO FAZER NAO ENTREGAR  -> ", ; Format(valorFRete, "Currency")
                                          Printer.Print
                                          Printer.Print "***CANCELADO***.", ; Format(TotalDaCompra, "Currency")

        End Select


Printer.Print " ";
Printer.Print "";
Printer.Print " ";
Printer.Print " -----------------------------------------------"
Printer.Print , ObservacoesparaEntrega
Printer.Print " OP > "; operador; Spc(2); "REG"; Datahora
Printer.Print " teste"
Printer.Print ; " ½"


Printer.EndDoc
                                          rs.Close
                                        

                            
                        
                        
                        
                        
                  






 Set rs = Nothing
End If

 Set rs = Nothing

End Sub

Private Sub Text9_Change()
If Text9.Text <> 0 Then
Cancelar = True
End If
End Sub
Public Sub CancelarPedido()
Dim numPedido As Integer
Dim NpTitulo As String
If Form404.Label18.Caption <> "Label18" Then
numPedido = Form404.Label18.Caption
'If Revalidar = True Then
'Printer.Print " -----------------------------------------------"
'Printer.Print "MODIFICACAO DO PEDIDO:  ", numPedido
'Printer.Print ")"
'Printer.Print ""
'Printer.Print ""
'Printer.Print " -----------------------------------------------"
'Printer.Print "PEDIDO MODIFICADO AGUARDE NOVA COMANDA"
'Printer.Print " -----------------------------------------------"
'Printer.Print "MODIFICACAO  DESTE PEDIDO ", numPedido
'Printer.Print " -----------------------------------------------"
'Printer.Print " "
'Revalidar = False
ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con
rs.CursorLocation = adUseClient


sql = "UPDATE `at_contadorDePedidos` SET `intPedido` = '" & Form404.Label17.Caption & "',`contador` = '0', `situacaoImpressao` = '3'  WHERE `at_contadorDePedidos`.`id` =  '" & numPedido & "'"
                            rs.Open sql
'rs.Close
sql = "DELETE FROM `at_itens` WHERE `fk_pedido` = '" & numPedido & "' ORDER BY `id` DESC"
 rs.Open sql
 sql = "DELETE  FROM `at_Cupon` WHERE `numPedido` = '" & numPedido & "' ORDER BY `numPedido` ASC"
  rs.Open sql
sql = " DELETE FROM `Pagamento` WHERE `NUmPedido` ='" & numPedido & "' ORDER BY `NUmPedido` ASC"
 rs.Open sql
sql = "UPDATE `at_contadorDePedidos` SET `contador` = '0' WHERE `at_contadorDePedidos`.`id` =  '" & numPedido & "'"
                            rs.Open sql

 Set rs = Nothing
'End If
'comander6
'imprimirCupon
Cancelar = True
Form401.Text26.Text = 1
'Form404.Text9.Text = 1
Unload Form404
Unload Form401
Unload Form402
Form2.Show
End If
End Sub



Public Sub verificarsepossuiaFrete()
If numerodopedido <> "0" Then
Form404.Adodc2.RecordSource = ""

Form404.Adodc2.CommandType = adCmdText

Form404.Adodc2.RecordSource = "SELECT `frete` FROM `at_frete` WHERE `fk_numPedido` ='" & numerodopedido & "'"

Form404.Adodc2.Refresh
Form500.ProgressBar1.Value = 60
If Form404.Adodc2.Recordset.BOF = False Then
Form407.Text7.Text = Replace(Form404.Adodc2.Recordset.Fields("frete").Value, ",", ".")
Else

            ConServer
            
            Dim sql As String
            Dim rs As New ADODB.Recordset
            Set rs = New ADODB.Recordset
            
            Set rs.ActiveConnection = con
            rs.CursorLocation = adUseClient
            
            
            sql = "INSERT INTO `at_frete` (`id`, `frete`, `fk_numPedido`) VALUES (NULL, '0', '" & numerodopedido & "') "
                                        rs.Open sql
            
            
             Set rs = Nothing




End If
End If
End Sub

Public Sub Atualize()

Form404.Adodc1.RecordSource = ""

Form404.Adodc1.CommandType = adCmdText

Form404.Adodc1.RecordSource = "SELECT `Quantidade`,`descrição`,`valor`,`id` FROM `at_itens` WHERE `fk_pedido` = '" & Label18.Caption & "'"

Form404.Adodc1.Refresh


Form404.DataGrid1.Columns(0).Caption = "Qtde"
Form404.DataGrid1.Columns(0).Width = 400
Form404.DataGrid1.Columns(1).Width = 4500
Form404.DataGrid1.Columns(2).Width = 1800
End Sub

Public Sub Carregarvaloresdeeditor()
repaceEnderecoManual
Form404.Label12.Caption = Form2.Text1.Text
Form404.Label17.Caption = Form2.DataCombo1.Text
Form404.Label14.Caption = Form2.Text2.Text
End Sub

Public Sub metSobrescrevaoCUpom()
ConServer
Dim numerodopedido As Integer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs.ActiveConnection = con

rs.CursorLocation = adUseClient

numerodopedido = Form404.Label18.Caption

sql = "UPDATE `robofi61_order_taker`.`at_Cupon` SET `nomeCliente` = ' " & Label12.Caption & " ', `endereco` =' " & Label13.Caption & " ', `telefone` = ' " & Label15.Caption & " ', `referencia` = ' " & Label12.Caption & " ', `loja` = ' " & Label17.Caption & " ' WHERE (`numPedido` = '" & numerodopedido & "')"
  rs.Open sql
 Set rs = Nothing


End Sub
