VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form50 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cardápio"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15675
   Icon            =   "Criecardapio.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   15675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Criecardapio.frx":0A02
      Height          =   1095
      Left            =   5400
      TabIndex        =   41
      Top             =   9000
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1931
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
   Begin VB.TextBox Text3 
      DataField       =   "idCardapio"
      DataSource      =   "Adodc3"
      Height          =   285
      Left            =   11160
      TabIndex        =   39
      Text            =   "Text3"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   11400
      TabIndex        =   28
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Index           =   1
      Left            =   360
      TabIndex        =   26
      Top             =   1320
      Width           =   10455
      Begin VB.CommandButton Command5 
         Caption         =   "Personalizar  botão"
         Height          =   615
         Left            =   600
         Picture         =   "Criecardapio.frx":0A17
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ingredientes"
         Height          =   4335
         Left            =   3600
         TabIndex        =   37
         Top             =   480
         Width           =   6135
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   495
            Left            =   240
            Top             =   3600
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   873
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
            RecordSource    =   "ingredientes_por_id_anexos"
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
         Begin MSDataListLib.DataList DataList1 
            Bindings        =   "Criecardapio.frx":1419
            Height          =   2940
            Left            =   480
            TabIndex        =   38
            Top             =   480
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   5186
            _Version        =   393216
            ListField       =   "ingredientes_por_id_anexoscol"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Preparo"
         Height          =   615
         Left            =   600
         Picture         =   "Criecardapio.frx":142E
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2280
         Width           =   2535
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Medidas "
         Height          =   615
         Left            =   600
         Picture         =   "Criecardapio.frx":1E30
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Inserir ingredientes"
         Height          =   615
         Left            =   600
         Picture         =   "Criecardapio.frx":2832
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Salvar Ingrediente"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3600
         Picture         =   "Criecardapio.frx":3234
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   5400
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Height          =   615
      Left            =   960
      Picture         =   "Criecardapio.frx":3C36
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveLast 
      Caption         =   "Mover"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6000
      Picture         =   "Criecardapio.frx":4638
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton CmdMoveFrist 
      Caption         =   "Mover"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      Picture         =   "Criecardapio.frx":503A
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5040
      Picture         =   "Criecardapio.frx":5A3C
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3120
      Picture         =   "Criecardapio.frx":643E
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar"
      CausesValidation=   0   'False
      Height          =   615
      Left            =   2040
      Picture         =   "Criecardapio.frx":6E40
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4080
      Picture         =   "Criecardapio.frx":7842
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   4440
      Top             =   8280
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
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
      RecordSource    =   "Cardapio"
      Caption         =   "Adodc3"
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
   Begin VB.Frame Frame1 
      Height          =   6255
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Top             =   1560
      Width           =   10455
      Begin VB.TextBox Text11 
         DataField       =   "codCategoria"
         DataSource      =   "Adodc3"
         Height          =   375
         Left            =   5760
         TabIndex        =   32
         Text            =   "Text11"
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         DataField       =   "tipo"
         DataSource      =   "Adodc3"
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
         Height          =   375
         Left            =   4080
         TabIndex        =   31
         Text            =   "10"
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox Text9 
         DataField       =   "codTitulo"
         DataSource      =   "Adodc3"
         Height          =   375
         Left            =   1920
         TabIndex        =   30
         Text            =   "Text9"
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         Picture         =   "Criecardapio.frx":8244
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Salvar e &Manter"
         Enabled         =   0   'False
         Height          =   855
         Left            =   9480
         Picture         =   "Criecardapio.frx":8C46
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Salvar N&ovo"
         Enabled         =   0   'False
         Height          =   735
         Left            =   9480
         Picture         =   "Criecardapio.frx":9648
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5520
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Height          =   3855
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   10335
         Begin VB.TextBox Text13 
            DataField       =   "codMedidas"
            DataSource      =   "Adodc3"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5640
            TabIndex        =   33
            Text            =   "Text13"
            Top             =   2280
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox Text12 
            DataField       =   "Medida"
            DataSource      =   "Adodc3"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3120
            TabIndex        =   4
            Text            =   "12"
            Top             =   3000
            Width           =   3015
         End
         Begin VB.TextBox Text5 
            DataField       =   "codigo"
            DataSource      =   "Adodc3"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            MaxLength       =   4
            TabIndex        =   2
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Height          =   375
            Left            =   120
            Picture         =   "Criecardapio.frx":A04A
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox Text2 
            DataField       =   "Descricao"
            DataSource      =   "Adodc3"
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
            Height          =   645
            Left            =   360
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   1  'Horizontal
            TabIndex        =   3
            Text            =   "Criecardapio.frx":AA4C
            Top             =   1800
            Width           =   9255
         End
         Begin VB.TextBox Text4 
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
            Height          =   375
            Left            =   6480
            TabIndex        =   5
            Top             =   3000
            Width           =   2655
         End
         Begin VB.Label Label9 
            Caption         =   "Código"
            Height          =   255
            Left            =   360
            TabIndex        =   29
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "  Inserir  itens no cardápio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            TabIndex        =   18
            Top             =   0
            Width           =   3255
         End
         Begin VB.Label Label4 
            Caption         =   "Descrição do Produto"
            Height          =   255
            Left            =   360
            TabIndex        =   17
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Valor"
            Height          =   255
            Left            =   6480
            TabIndex        =   16
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Medidas Possíveis "
            Height          =   255
            Left            =   3120
            TabIndex        =   15
            Top             =   2760
            Width           =   1575
         End
      End
      Begin VB.TextBox Text1 
         DataField       =   "Titulo"
         DataSource      =   "Adodc3"
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
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Text            =   "1"
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Categoria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   7095
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   12515
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Cadastro "
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Ingrediente"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label40 
      Caption         =   "Crie um cardápio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      TabIndex        =   9
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form50"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
 
Dim novoCadastro As Boolean
Dim TipOs, titulo, Medida, descricao As String
Dim cmandEditar As Integer

Dim valor As Double





Private Sub C_Click()

End Sub

Private Sub CmdConsultar_Click()
limpezaSalvamentoVazios
novoCadastro = False
Form52.Show
cmdEditar.Enabled = True
Form52.Text1.Text = "50"
'ipossibilitar alteração de tabela modo de execulsao
'ATIVAR AS MOBILIDADES CMD MOVE LAST MOVE FRIST
cmdMoveLast.Enabled = True
CmdMoveFrist.Enabled = True


End Sub

Private Sub cmdEditar_Click()
novoCadastro = False
Text1.Enabled = True
Text4.Enabled = True
Text2.Enabled = True
Text5.Enabled = True
Text10.Enabled = True
Text12.Enabled = True
    Command1.Enabled = True
    cmdSalvar.Enabled = True
    Command2.Enabled = True
cmandEditar = 1


Text1.SetFocus
RefatoricCodesBrid
End Sub

Private Sub CmdExcluir_Click()
novoCadastro = False
Dim resp As Integer
resp = MsgBox("tem certeza que deseja escluir este produto", vbYesNo, "Excluir ?")
If resp = 6 Then
'CommonDialog1.CancelError = True
'trarar erro
On Error GoTo error
Adodc3.Recordset.Delete
Adodc3.Refresh
'Form52.Adodc1.Refresh
'Form52.Adodc2.Refresh
'Form52.Adodc3.Refresh
End If
Exit Sub

error:
MsgBox "O produto já foi excluido !<Sincronizando banco de dados >", , "Produto Excluido"
'Form52.Adodc1.Refresh
'Form52.Adodc2.Refresh
'Form52.Adodc3.Refresh
Exit Sub



End Sub

Private Sub CmdMoveFrist_Click()
novoCadastro = False

'trarar erro
On Error GoTo error
Adodc3.Recordset.MovePrevious
REVOLTARlistaAdodc2itensjaescolhidos
cmdMoveLast.Enabled = True
Exit Sub

error:

CmdMoveFrist.Enabled = False
Exit Sub




End Sub

Private Sub cmdMoveLast_Click()
novoCadastro = False


'trarar erro
On Error GoTo error
Adodc3.Recordset.MoveNext
REVOLTARlistaAdodc2itensjaescolhidos
CmdMoveFrist.Enabled = True
Exit Sub

error:
cmdMoveLast.Enabled = False
Exit Sub

End Sub
Public Sub REVOLTARlistaAdodc2itensjaescolhidos()

Adodc1.RecordSource = ""

Adodc1.CommandType = adCmdText
                       
Adodc1.RecordSource = "SELECT * FROM `ingredientes_por_id_anexos` WHERE `chaveDaProdutp` = '" & Form50.Text3.Text & "'"
Adodc1.Refresh
End Sub

Private Sub cmdNovo_Click()
ajusta_container
Text1.Enabled = True
Text4.Enabled = True
Text2.Enabled = True
Text5.Enabled = True
Text10.Enabled = True
Text12.Enabled = True
novoCadastro = True


Command1.Enabled = True

Command4.Enabled = True
cmdSalvar.Enabled = True
Command2.Enabled = True

Adodc3.Recordset.AddNew

Text1.SetFocus

End Sub

Private Sub cmdSalvar_Click()

SalvarEnovo

cmdSalvar.Enabled = False
cmdNovo.Enabled = True

Command1.Enabled = True
cmdSalvar.Enabled = True
RefatoricCodesSolids
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
If novoCadastro = True Then
  SalvarEmanter
  'novoCadastro = False
End If




If cmandEditar = 1 Then
     
    SalvarEmanter
    Command1.Enabled = True
    cmdSalvar.Enabled = True
 Else
 If Text5.Text <> "" Then
    If verificarExclusividadeDoCOdigo(Text5.Text) Then
    cmandEditar = 0
    SalvarEmanter
    Command1.Enabled = True
    cmdSalvar.Enabled = True
   
    Else
            If cmandEditar <> 1 Then
            MsgBox "Codigo :" & Text5.Text & " Já foi inserido em outro item", vbYes, "Codigo repitido!!"
            
            Text5.Text = ""
            Text5.SetFocus
            Else
            MsgBox "Produto Codigo : " & Text5.Text & "Editado com sucesso! ", vbYes, "Editado!"
            End If
            End If
    End If
End If
Command1.Enabled = True
RefatoricCodesSolids

End Sub

Private Sub Command11_Click()
REVOLTARlistaAdodc2itensjaescolhidos
End Sub

Private Sub Command10_Click()
   
Form51_7.Show
Form51_7.Caption = "Informe a instruçao de prepaaro  para " & Form50.Text1.Text & " Produto " & Form50.Text2.Text & "  medida " & Form50.Text12.Text & " "
nivelarcomOsinteressesmemor

End Sub

Private Sub Command2_Click()
novoCadastro = False
Command2.Enabled = False
Salvar

End Sub

Private Sub Command3_Click()



MsgBox "Mantenha pressionado o botão ALT e digite 171 para criar o caractere => 1/2 ", , "Dica"
        
End Sub

Private Sub Command4_Click()
Form51.Show
End Sub

Private Sub DataCombo1_LostFocus()
'DataCombo1.Text = Format(DataCombo1.Text, ">")
'Medida = DataCombo1.Text
End Sub

Private Sub DataCombo2_LostFocus()
'DataCombo2.Text = Format(DataCombo2.Text, ">")
'TipOs = DataCombo2.Text
'If DataCombo3.Text <> "" Then
'Text1.Text = Format(DataCombo3.Text, ">")
'End If
'DataCombo3.Visible = False

End Sub

Private Sub DataCombo3_DblClick(Area As Integer)
'Text1.Text = Format(DataCombo3.Text, ">")
VoltaraoNormal
'DataCombo3.Visible = False
'DataCombo2.SetFocus



End Sub

Private Sub DataCombo3_LostFocus()
'Text1.Text = Format(DataCombo3.Text, ">")
VoltaraoNormal
'DataCombo3.Visible = False
'DataCombo2.SetFocus

End Sub

Private Sub Command6_Click()



'ConServer

'Dim sql As String
'Dim rs As New ADODB.Recordset

'Set rs = New ADODB.Recordset
'Set rs.ActiveConnection = con
'
'sql = "SELECT * FROM `coordenadasgeonow` WHERE `Latitude` BETWEEN " & latidudinalMinimo & " AND " & latitudinalMaximos & " AND `Longitude` BETWEEN " & LongitudeMinima & "  AND  " & LongitudeMaxima & "  ORDER BY `idcoordenadasGeoNow` ASC"

'rs.Open sql




'abre a conexao
'rs.Open sql$, adocn, adOpenStatic
'define a fonte de dados para a conexao ativa
'Set DataGrid1.DataSource = rs

'libera a conexao
' Set rs = Nothing






'If Adodc1.Recordset.EOF = False Then
'Adodc1.Refresh

'rs.Close sql
'i dentificadorIndicesTabela
'SituaçãoPrimeiraLinha

'End If
End Sub

Private Sub Command5_Click()
Form444.Show
Form50.Hide
End Sub

Private Sub Command8_Click()
Form51_4.Caption = "Crie ingredientes para " & Form50.Text1.Text & " Produto " & Form50.Text2.Text & "  medida " & Form50.Text12.Text & " "
Form51_4.Show

End Sub

Private Sub Command9_Click()
Form51_5.Show
Form51_5.Caption = "Crie a receita  para " & Form50.Text1.Text & " Produto " & Form50.Text2.Text & "  medida " & Form50.Text12.Text & " "
nivelarcomOsinteresses

End Sub

Private Sub DataList1_Click()
Call Command8_Click
End Sub

Private Sub ajusta_container()
Dim i As Integer
With TabStrip1
For i = 1 To .Tabs.Count
Frame1(i - 1).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
Next
End With

TabStrip1.Tabs(1).Selected = True
   TabStrip1_Click
End Sub

Private Sub Form_Load()
 
Dim contemax As Integer
contemax = Len("COMBO ESPECIAL : 2  X-TUDO ESPECIAL + 2 FRITAS 200GR + 2 PITCHULINHA")
 ajusta_container
 Adodc3.Recordset.MoveLast
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
limpezaSalvamentoVazios
Unload Form51
Unload Form51_3
Unload Form51_2
End Sub

Private Sub TabStrip1_Click()
Dim i As Integer


'trarar erro
On Error GoTo error



i = TabStrip1.SelectedItem.Index

Frame1(i - 1).ZOrder

  
If i = 2 Then
CmdMoveFrist.Enabled = True
cmdMoveLast.Enabled = True
End If
REVOLTARlistaAdodc2itensjaescolhidos
'Adodc3.Refresh
'Adodc3.Recordset.MoveLast


Exit Sub

error:
'Adodc3.RecordSource = ""

'Adodc3.CommandType = adCmdText
'
'Adodc3.RecordSource = "SELECT * FROM `Cardapio` ORDER BY `idCardapio` ASC"

'Adodc3.Refresh

'Adodc3.Recordset.MoveLast


Exit Sub
End Sub

Private Sub Text1_DblClick()
Text1.Text = ""
Form51.Show

End Sub

Private Sub Text1_GotFocus()
If Text1.Text = "" Then
Form51.Show
End If
End Sub

Private Sub Text10_DblClick()
Text10.Text = ""
Form51_2.Show

End Sub

Private Sub Text10_GotFocus()
If Text10.Text = "" Then
Form51_2.Show
End If
End Sub

Private Sub Text12_DblClick()
Text12.Text = ""
Form51_3.Show
'Form51_3.DataCombo1.SelText
'Form51_3.DataCombo1.SetFocus

End Sub

Private Sub Text12_GotFocus()
If Text12.Text = "" Then
Form51_3.Show
End If
End Sub

Private Sub Text12_LostFocus()
Text12.Text = Format(Text12.Text, ">")
End Sub

Private Sub Text2_DblClick()
Text2.Text = ""
End Sub
Private Sub Text42_DblClick()
Text4.Text = ""
End Sub

Private Sub Text2_LostFocus()
Text2.Text = Format(Text2.Text, ">")
descricao = Text2.Text
End Sub

Private Sub Text4_DblClick()
Text4.Text = ""
End Sub

Private Sub Text4_LostFocus()
If Text4.Text <> "" And Text4.Enabled = True Then
Dim testo4 As Double
    Text4.Text = Replace(Text4.Text, ",", ".")
 
testo4 = CDbl(Text4.Text)
'Text4.Text = Format(Text4.Text, "Currency")
End If



End Sub



Public Function VerificaSalvamento() As Boolean
If Text4.Text <> "" And Text2.Text <> "" And Text12.Text <> "" And Text1.Text <> "" And Text10.Text <> "" Then
VerificaSalvamento = True
Else
    If Text4.Text = "" Then
    MsgBox "Insira um valor para o produto", vbInformation, "Valor do Produto?"
    Text4.SetFocus
    End If
    If Text2.Text = "" Then
    MsgBox "Informe o produto em descrição", vbInformation, "Descrição do Produto?"
    Text2.SetFocus
    End If
    If Text12.Text = "" Then
    MsgBox "Informe ou selecione uma medida para o produto", vbInformation, "Medida?"
    Text12.SetFocus
    End If
    If Text1.Text = "" Then
    MsgBox "Informe um titulo para criar o cardápio ", vbInformation, "Titulo?"
    Text2.SetFocus
    End If
    If Text10.Text = "" Then
    MsgBox "Informe ou selecione a melhor categoria, para seu novo produto", vbInformation, "Descrição do Produto?"
    Text10.SetFocus
    End If
VerificaSalvamento = False
End If


End Function

Public Sub SalvarEmanter()
If (VerificaSalvamento) Then
    'ConsulteSemelhancaNOtipo (TipOs)
    'ConsulteSemelhancaMedida (Medida)
    titulo = Text1.Text
    Medida = Text12.Text
    If Text12.Text = "" Then
    MsgBox "É obrigatório informar a medida", , "Medida?"
    Text12.SetFocus
    Exit Sub
    End If
    TipOs = Text10.Text
    descricao = Text2.Text
    Text4.Text = DBLcurre(Text4.Text)
    Adodc3.Recordset("valor").Value = Text4.Text
   'trarar erro
   On Error GoTo error
        Adodc3.Recordset.Update
        Adodc3.Recordset.MoveLast
           ' Adodc3.Refresh
        Adodc3.Recordset.AddNew
        Text1.Text = titulo
        'DataCombo1.Text = meDida
        Text10.Text = TipOs
        Text2.Text = descricao
        Text4.Text = ""
        Command1.Enabled = False
        
        MsgBox "Salvo com sucesso!!!", vbYes, "Iten inserido!"
        Command1.Enabled = True
        
        Text5.SetFocus
    
        End If
        Command1.Enabled = True
        Exit Sub


error:
'usage
'Sleep 3000

'   Adodc3.Recordset.Update
'        Adodc3.Recordset.MoveLast
           Adodc3.Refresh
        Adodc3.Recordset.AddNew
        Text1.Text = titulo
        'DataCombo1.Text = meDida
        Text10.Text = TipOs
        Text2.Text = descricao
        Text4.Text = ""
        Command1.Enabled = False
        
        MsgBox "Salvo com sucesso!!!", vbYes, "Iten inserido!"
        Command1.Enabled = True
        
        Text5.SetFocus
    
        'End If
        Command1.Enabled = True
        Exit Sub

Exit Sub

        
        
        
        
        
        
End Sub

Public Sub criarsemrepeticaoosTitulos()

Adodc3.RecordSource = ""

Adodc3.CommandType = adCmdText

Adodc3.RecordSource = "SELECT DISTINCT`Titulo` FROM `Cardapio` "

Adodc3.Refresh

 
End Sub


Public Sub VoltaraoNormal()

Adodc3.RecordSource = ""

Adodc3.CommandType = adCmdText

Adodc3.RecordSource = "SELECT * FROM `Cardapio` "




titulo = Text1.Text
Adodc3.Refresh
Adodc3.Recordset.AddNew
Text1.Text = titulo


 
End Sub

Public Sub SalvarEnovo()
If (VerificaSalvamento) Then
'ConsulteSemelhancaNOtipo (TipOs)
'ConsulteSemelhancaMedida (Medida)
titulo = Text1.Text
Medida = Text12.Text
TipOs = Text10.Text
descricao = Text2.Text
  Text4.Text = DBLcurre(Text4.Text)
    Adodc3.Recordset("valor").Value = Text4.Text
 'trarar erro
   On Error GoTo error

    Adodc3.Recordset.Update
    Adodc3.Recordset.MoveLast
    'Adodc3.Refresh
    Adodc3.Recordset.AddNew
    Text1.Text = titulo
    'DataCombo1.Text = meDida
    Text10.Text = TipOs
    'Text2.Text = descriCao
    Text4.Text = ""
    MsgBox "Salvo com sucesso!!!", vbYes, "Iten inserido!"
    cmdSalvar.Enabled = True
    
   Text5.SetFocus
    End If
          
          
                  Exit Sub


error:
  'usage
'Sleep 3000

    Adodc3.Recordset.Update
    Adodc3.Recordset.MoveLast
    'Adodc3.Refresh
    Adodc3.Recordset.AddNew
    Text1.Text = titulo
    'DataCombo1.Text = meDida
    Text10.Text = TipOs
    'Text2.Text = descriCao
    Text4.Text = ""
    MsgBox "Salvo com sucesso!!!", vbYes, "Iten inserido!"
    cmdSalvar.Enabled = True
    
   Text5.SetFocus
 
    
        Exit Sub

Exit Sub

        
        
        

          
          
End Sub



Public Sub Salvar()

cmdNovo.Enabled = True
If (VerificaSalvamento) Then
'ConsulteSemelhancaNOtipo (TipOs)
'ConsulteSemelhancaMedida (Medida)
 Text4.Text = DBLcurre(Text4.Text)
Adodc3.Recordset("valor").Value = Text4.Text

    'trarar erro
   On Error GoTo error

'Adodc3.Recordset("valor").Value = Text4.Text

    Adodc3.Recordset.Update
    Adodc3.Recordset.MoveFirst
    Adodc3.Refresh
    
    
    MsgBox "Salvo com sucesso!!!", vbYes, "Cardápio Salvo!"
    nivelOenabled
    Adodc3.Recordset.MoveLast
    End If
          
                  Exit Sub


error:
  'usage
'Sleep 3000
'causes program to pause for 3 seconds
'    Adodc3.Recordset.Update
 '   Adodc3.Recordset.MoveFirst
  '  Adodc3.Refresh
    
    
    MsgBox "Salvo com sucesso!!!", vbYes, "Cardápio Salvo!"
    nivelOenabled
'    Adodc3.Recordset.MoveLast
    
    MsgBox "Não salvo erro!!!", vbYes, "Não  Salvo!"
  
     
    Exit Sub

Exit Sub



Adodc3.Recordset.MoveLast

End Sub


Public Sub nivelOenabled()
  Command1.Enabled = False
    Command2.Enabled = False
      Command4.Enabled = False
        cmdSalvar.Enabled = False
            cmdEditar.Enabled = False
                CmdExcluir.Enabled = False
                    CmdMoveFrist.Enabled = False
                        cmdMoveLast.Enabled = False
                        
        
 Text1.Enabled = False
    Text2.Enabled = False
        Text4.Enabled = False
 Text10.Enabled = False
 Text12.Enabled = False
 Text5.Enabled = False
End Sub

Private Sub Text5_DblClick()
Text5.Text = ""
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = (SoNumeros(KeyAscii))
    If KeyAscii = 0 Then
    End If
End Sub

Private Sub Text7_Change()
If Text7.Text <> "" Then
RetornoDeconsulta (Text7.Text)
End If

End Sub
Public Sub RetornoDeconsulta(indice As Integer)

Adodc3.RecordSource = ""

Adodc3.CommandType = adCmdText

Adodc3.RecordSource = "SELECT * FROM `Cardapio` WHERE `idCardapio` = ' " & indice & " '"

Adodc3.Refresh
Text4.Text = Format(Adodc3.Recordset("valor").Value, "Currency")
 
End Sub

Public Function verificarExclusividadeDoCOdigo(codigo As Integer) As Boolean

ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

sql = "SELECT * FROM `Cardapio` WHERE `codigo` = ' " & codigo & " '"
rs.Open sql



If rs.BOF = False Then
verificarExclusividadeDoCOdigo = False
Else
verificarExclusividadeDoCOdigo = True


End If
 Set rs = Nothing


 









End Function

Public Sub RefatoricCodesBrid()
'exemple conectedd


Dim i As Integer


'enviar
Dim nomeTitulo As String
Dim nomeCategoria As String
Dim nomeMedida As String
'receber
Dim IdTituloVar As Integer
Dim iDCategoriaVar As Integer
Dim iDmedidaVar As Integer
'text
nomeTitulo = Text1.Text
nomeCategoria = Text10.Text
nomeMedida = Text12.Text



ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

sql = "SELECT * FROM `Cardapio_Tipo`,`Cardapio_medidas`,`TituloCardapio` WHERE `tipos` LIKE '" & nomeCategoria & "'AND `unidade` LIKE '" & nomeMedida & "' AND `nometitulo` LIKE '" & nomeTitulo & "'"
rs.Open sql



If rs.BOF = False Then
IdTituloVar = rs.Fields("idTituloCardapio").Value
iDCategoriaVar = rs.Fields("idCardapio_Tipo").Value
iDmedidaVar = rs.Fields("idCardapio_medidas").Value
End If
 Set rs = Nothing



Text9.Text = IdTituloVar
Text11.Text = iDCategoriaVar
Text13.Text = iDmedidaVar




End Sub

Public Sub RefatoricCodesSolids()
'exemple conectedd


'enviar
Dim nomeTitulo As String
Dim nomeCategoria As String

'receber
Dim IdTituloVar As Integer
Dim iDCategoriaVar As Integer

'text
nomeTitulo = Text1.Text
nomeCategoria = Text10.Text



ConServer

Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

sql = "SELECT * FROM `Cardapio_Tipo`,`TituloCardapio` WHERE `tipos` LIKE '" & nomeCategoria & "' AND `nometitulo` LIKE '" & nomeTitulo & "'"
rs.Open sql



If rs.BOF = False Then
IdTituloVar = rs.Fields("idTituloCardapio").Value
iDCategoriaVar = rs.Fields("idCardapio_Tipo").Value

End If
 Set rs = Nothing



Text9.Text = IdTituloVar
Text11.Text = iDCategoriaVar





End Sub


Public Sub nivelarcomOsinteresses()
Dim indice As Integer
indice = Text3.Text

Form51_5.Adodc1.RecordSource = ""

Form51_5.Adodc1.CommandType = adCmdText

Form51_5.Adodc1.RecordSource = "SELECT * FROM `ingredientes_por_id_anexos` WHERE `chaveDaProdutp` =  ' " & indice & " '"

Form51_5.Adodc1.Refresh
'Text4.Text = Format(Adodc3.Recordset("valor").Value, "Currency")

End Sub
Public Sub nivelarcomOsinteressesmemor()
Dim indice As Integer
indice = Text3.Text

Form51_7.Adodc1.RecordSource = ""

Form51_7.Adodc1.CommandType = adCmdText

Form51_7.Adodc1.RecordSource = "SELECT * FROM `ingredientes_por_id_anexos` WHERE `chaveDaProdutp` =  ' " & indice & " '"

Form51_7.Adodc1.Refresh
'Text4.Text = Format(Adodc3.Recordset("valor").Value, "Currency")

Form51_7.Text2.Text = Text3.Text
'pesquiseatributosParateemarket (indice)

'movimetar tambem o testo do atendimetno
Form51_7.Adodc3.RecordSource = ""

Form51_7.Adodc3.CommandType = adCmdText
                       
Form51_7.Adodc3.RecordSource = "SELECT * FROM `ProntoAtendimentoTelemarket` WHERE `NumeroPrato` =  '" & indice & "' "
Form51_7.Adodc3.Refresh


If (Form51_7.Adodc3.Recordset.BOF = False) Then
    'GIRAR O ADO PARA A POSICAO DO PRATO
    
    Form51_7.Check1.Value = 1
    Form51_7.Text5.Enabled = True
    Form51_7.Text5.SetFocus
    
       
        
    Else
    Form51_7.Check1.Value = 0
     Form51_7.Text5.Text = ""
    Form51_7.Text5.Enabled = False
   'form51_7.Adodc3.Recordset.AddNew
   
    
   End If






End Sub

Public Sub pesquiseatributosParateemarket(indice)
Dim NumContador As Integer


ConServer

Dim sql As String
Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

sql = "SELECT * FROM `ProntoAtendimentoTelemarket` WHERE `NumeroPrato` =  '" & indice & "' "

rs.Open sql


    If (rs.EOF = False) Then
    'GIRAR O ADO PARA A POSICAO DO PRATO
    GIREPARAPOSICAOdopratotelemaaarketing (indice)
    
    Form51_7.Check1.Value = 1
    Form51_7.Text5.Enabled = True
    Form51_7.Text5.SetFocus
    
       
        
    Else
     Form51_7.Text5.Text = ""
    Form51_7.Text5.Enabled = False
   'form51_7.Adodc3.Recordset.AddNew
   
    
   End If

   Set rs = Nothing



End Sub
Public Sub GIREPARAPOSICAOdopratotelemaaarketing(indice As Integer)

Form51_7.Adodc3.RecordSource = ""

Form51_7.Adodc3.CommandType = adCmdText
                       
Form51_7.Adodc3.RecordSource = "SELECT * FROM `ProntoAtendimentoTelemarket` WHERE `NumeroPrato` =  '" & indice & "' "
Form51_7.Adodc3.Refresh

 
End Sub

Public Sub limpezaSalvamentoVazios()
ConServer

Dim sql As String
Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = con

sql = "DELETE  FROM `Cardapio` WHERE `valor` = 0 "

rs.Open sql


sql = "DELETE  FROM `Cardapio` WHERE `Descricao` LIKE '' "
rs.Open sql

 Set rs = Nothing





End Sub
