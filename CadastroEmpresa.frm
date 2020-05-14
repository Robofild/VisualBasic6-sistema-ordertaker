VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form11 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro da Empresa"
   ClientHeight    =   6540
   ClientLeft      =   2625
   ClientTop       =   1050
   ClientWidth     =   19485
   Icon            =   "CadastroEmpresa.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   19485
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox CmbFormatacao 
      Height          =   315
      ItemData        =   "CadastroEmpresa.frx":0A02
      Left            =   7320
      List            =   "CadastroEmpresa.frx":0A0F
      TabIndex        =   90
      Text            =   "Formatação "
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Index           =   2
      Left            =   9720
      TabIndex        =   73
      Top             =   2160
      Width           =   9495
      Begin VB.TextBox TxtreciboCnpjrecibo 
         DataField       =   "recibo_cnpj"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5520
         TabIndex        =   20
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox txtRefernciarecibo 
         DataField       =   "recibo_referencia"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   360
         TabIndex        =   19
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox Txtendercorecibo 
         DataField       =   "recibo_nome"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   16
         Top             =   480
         Width           =   4935
      End
      Begin VB.TextBox tXTbAIRROrecibo 
         DataField       =   "recibo_bairro"
         DataSource      =   "Adodc1"
         Height          =   405
         Index           =   1
         Left            =   360
         TabIndex        =   21
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox TXTcIDADErecibo 
         DataField       =   "Recibo_cidade"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   22
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox tXTcEPrecibo 
         DataField       =   "recibo_cep"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   23
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox TXTuFrecibo 
         DataField       =   "recibo_uf"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   1
         Left            =   6480
         TabIndex        =   89
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox tXTcONTATOrecibo 
         DataField       =   "recibo_contato"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   1
         Left            =   7200
         TabIndex        =   24
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox numerorecibo 
         DataField       =   "recibo_numero"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   17
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox textComplementoRecibo 
         DataField       =   "recibo_complemento"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   1
         Left            =   6480
         TabIndex        =   18
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox tXTtELEFONErecibo 
         DataField       =   "recibo_telefone"
         DataSource      =   "Adodc1"
         Height          =   405
         Index           =   1
         Left            =   360
         TabIndex        =   25
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox tXTfAXrecibo 
         DataField       =   "recibo_fax"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   1
         Left            =   2280
         TabIndex        =   26
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox tXTcELULARrecibo 
         DataField       =   "recibo_celular"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   27
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox tXTsITErecibo 
         DataField       =   "recibo_site"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   1
         Left            =   5880
         TabIndex        =   28
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox tEXTeMAILrecibo 
         DataField       =   "recibo_email"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   1
         Left            =   7440
         TabIndex        =   29
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "CPF/CNPJ"
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
         Left            =   5520
         TabIndex        =   88
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label16 
         Caption         =   "Referência"
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
         Left            =   360
         TabIndex        =   87
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Endereço"
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
         Index           =   1
         Left            =   360
         TabIndex        =   86
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   85
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Complemento"
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
         Index           =   1
         Left            =   6480
         TabIndex        =   84
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Bairro "
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
         Index           =   1
         Left            =   360
         TabIndex        =   83
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Cidade"
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
         Index           =   1
         Left            =   2520
         TabIndex        =   82
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "UF"
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
         Index           =   1
         Left            =   6480
         TabIndex        =   81
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "CEP "
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
         Index           =   1
         Left            =   4680
         TabIndex        =   80
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Contato"
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
         Index           =   1
         Left            =   7200
         TabIndex        =   79
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Telefone"
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
         Index           =   1
         Left            =   360
         TabIndex        =   78
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Fax"
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
         Index           =   1
         Left            =   2280
         TabIndex        =   77
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Site"
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
         Index           =   1
         Left            =   5880
         TabIndex        =   76
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Celular"
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
         Index           =   1
         Left            =   4200
         TabIndex        =   75
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Email"
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
         Index           =   1
         Left            =   7440
         TabIndex        =   74
         Top             =   2640
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Index           =   1
      Left            =   10560
      TabIndex        =   60
      Top             =   0
      Width           =   8775
      Begin MSDataListLib.DataCombo DTCBrEGIMEtRIBUTARIO 
         DataField       =   "regime_tributario"
         DataSource      =   "Adodc1"
         Height          =   765
         Left            =   360
         TabIndex        =   34
         Top             =   2280
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1349
         _Version        =   393216
         Style           =   1
         Text            =   ""
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dados do Contador "
         Height          =   1935
         Left            =   3000
         TabIndex        =   66
         Top             =   1320
         Width           =   5655
         Begin VB.TextBox tEXTTELFONEcont 
            DataField       =   "telefone_contador"
            DataSource      =   "Adodc1"
            Height          =   405
            Index           =   1
            Left            =   2880
            TabIndex        =   38
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox tXTeMAILCONT 
            DataField       =   "Email_contador"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox Text16 
            DataField       =   "cpf_cnpj_contador"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   2880
            TabIndex        =   36
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox TXTnoMEcONT 
            DataField       =   "nome_contador"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   240
            TabIndex        =   35
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label18 
            Caption         =   "Telefone"
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
            Left            =   2880
            TabIndex        =   72
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "Email"
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
            Left            =   240
            TabIndex        =   69
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label TXTcPFcONT 
            AutoSize        =   -1  'True
            Caption         =   "CPF/CNPJ"
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
            Left            =   2760
            TabIndex        =   68
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label14 
            Caption         =   "Nome"
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
            Left            =   240
            TabIndex        =   67
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.TextBox TXTcPF 
         DataField       =   "cnpj_cpf"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   360
         TabIndex        =   30
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox tXTcNAE 
         DataField       =   "cnae"
         DataSource      =   "Adodc1"
         Height          =   405
         Index           =   1
         Left            =   360
         TabIndex        =   33
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox tXtiDENTIDADE 
         DataField       =   "insc_estadual_identidade"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   1
         Left            =   2880
         TabIndex        =   31
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox tEXTmUNICIPAL 
         DataField       =   "insc_municipal_identidade"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   1
         Left            =   5880
         TabIndex        =   32
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "CPF/CNPJ"
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
         Left            =   360
         TabIndex        =   65
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "INSC.ESTADUAL/IDENTIDADE"
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
         Left            =   2880
         TabIndex        =   64
         Top             =   480
         Width           =   2715
      End
      Begin VB.Label Label24 
         Caption         =   "INSC.Municipal"
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
         Left            =   5880
         TabIndex        =   63
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "CNAE"
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
         Left            =   360
         TabIndex        =   62
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "Regime Tributário"
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
         Left            =   360
         TabIndex        =   61
         Top             =   1920
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Index           =   0
      Left            =   600
      TabIndex        =   46
      Top             =   2160
      Width           =   9375
      Begin VB.TextBox textfantazia 
         DataField       =   "nome_fantazia"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   2
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtrazao 
         DataField       =   "razao_social"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox txtemail 
         DataField       =   "email"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   0
         Left            =   7440
         TabIndex        =   15
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox tXTsITE 
         DataField       =   "site"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   0
         Left            =   5880
         TabIndex        =   14
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox tXTcELULAR 
         DataField       =   "celular"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   13
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox tXTfAX 
         DataField       =   "fax"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   0
         Left            =   2280
         TabIndex        =   12
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox tXTtELEFONE 
         DataField       =   "telefone"
         DataSource      =   "Adodc1"
         Height          =   405
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox TxtComplemento 
         DataField       =   "complemento"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   0
         Left            =   6480
         TabIndex        =   5
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtN 
         DataField       =   "numero"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   4
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox tXTcONTATO 
         DataField       =   "contato"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   0
         Left            =   7200
         TabIndex        =   10
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox TXTuF 
         DataField       =   "uf"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   0
         Left            =   6480
         TabIndex        =   9
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox tXTcEP 
         DataField       =   "cep"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   0
         Left            =   4680
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox TXTcIDADE 
         DataField       =   "cidade"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   7
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox tXTbAIRRO 
         DataField       =   "bairro"
         DataSource      =   "Adodc1"
         Height          =   405
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox TxtEndereco 
         DataField       =   "enderco"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   4935
      End
      Begin VB.Label Label20 
         Caption         =   "Nome Fantazia"
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
         Index           =   0
         Left            =   5520
         TabIndex        =   71
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Razão Social"
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
         Index           =   0
         Left            =   360
         TabIndex        =   70
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Email"
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
         Index           =   0
         Left            =   7440
         TabIndex        =   59
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Celular"
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
         Index           =   0
         Left            =   4200
         TabIndex        =   58
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Site"
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
         Index           =   0
         Left            =   5880
         TabIndex        =   57
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Fax"
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
         Index           =   0
         Left            =   2280
         TabIndex        =   56
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Telefone"
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
         Index           =   0
         Left            =   360
         TabIndex        =   55
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Contato"
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
         Index           =   0
         Left            =   7200
         TabIndex        =   54
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "CEP "
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
         Index           =   0
         Left            =   4680
         TabIndex        =   53
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "UF"
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
         Index           =   0
         Left            =   6480
         TabIndex        =   52
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Cidade"
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
         Index           =   0
         Left            =   2520
         TabIndex        =   51
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Bairro "
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
         Index           =   0
         Left            =   360
         TabIndex        =   50
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Complemento"
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
         Index           =   0
         Left            =   6480
         TabIndex        =   49
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   48
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Endereço"
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
         Index           =   0
         Left            =   360
         TabIndex        =   47
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4320
      Picture         =   "CadastroEmpresa.frx":0A3A
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      Height          =   615
      Left            =   1200
      Picture         =   "CadastroEmpresa.frx":143C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar"
      CausesValidation=   0   'False
      Height          =   615
      Left            =   2280
      Picture         =   "CadastroEmpresa.frx":1E3E
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3360
      Picture         =   "CadastroEmpresa.frx":2840
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5280
      Picture         =   "CadastroEmpresa.frx":3242
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton CmdMoveFrist 
      Caption         =   "Mover"
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      Picture         =   "CadastroEmpresa.frx":3C44
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdMoveLast 
      Caption         =   "Mover"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6240
      Picture         =   "CadastroEmpresa.frx":4646
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   360
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7440
      Top             =   840
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      RecordSource    =   "cadastro_da_empresa"
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
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4335
      Left            =   360
      TabIndex        =   45
      Top             =   1560
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7646
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Dados Cadastrais"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   6
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Dados Fiscais"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Texto Para Recibos"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7800
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   15
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadastroEmpresa.frx":5048
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadastroEmpresa.frx":5222
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadastroEmpresa.frx":53FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadastroEmpresa.frx":55D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadastroEmpresa.frx":57B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadastroEmpresa.frx":598A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadastroEmpresa.frx":5B64
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadastroEmpresa.frx":5D3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadastroEmpresa.frx":5F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadastroEmpresa.frx":60F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadastroEmpresa.frx":62CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadastroEmpresa.frx":64A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadastroEmpresa.frx":6680
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadastroEmpresa.frx":685A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CadastroEmpresa.frx":6A34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   7080
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNovo_Click()
Adodc1.Recordset.AddNew
txtrazao(0).SetFocus

End Sub

Private Sub Combo1_Change()
controleLetrass

End Sub

Private Sub Form_Load()
 ajusta_container
End Sub

Private Sub TabStrip1_Click()
Dim i As Integer

i = TabStrip1.SelectedItem.Index

Frame1(i - 1).ZOrder

  

End Sub
Private Sub ajusta_container()
Dim i As Integer
With TabStrip1
For i = 1 To .Tabs.Count
Frame1(i - 1).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
Next
End With

TabStrip1.Tabs(1).Selected = True
   
       
    
End Sub



Private Sub textfantazia_Change(Index As Integer)
AddoutRecibo
End Sub

Private Sub tXTbAIRRO_Change(Index As Integer)
AddoutRecibo
End Sub

Private Sub tXTcELULAR_Change(Index As Integer)
AddoutRecibo
End Sub

Private Sub tXTcEP_Change(Index As Integer)
AddoutRecibo
End Sub

Private Sub TXTcIDADE_Change(Index As Integer)
AddoutRecibo
End Sub

Private Sub TxtComplemento_Change(Index As Integer)
AddoutRecibo
End Sub

Private Sub tXTcONTATO_Change(Index As Integer)
AddoutRecibo
End Sub

Private Sub TXTcPF_Change()
AddoutRecibo
End Sub



Private Sub txtemail_Change(Index As Integer)
AddoutRecibo
End Sub

Private Sub TxtEndereco_Change(Index As Integer)
AddoutRecibo
End Sub

Private Sub tXTfAX_Change(Index As Integer)
AddoutRecibo
End Sub

Private Sub txtN_Change(Index As Integer)
AddoutRecibo
End Sub

Private Sub txtrazao_Change(Index As Integer)
AddoutRecibo

End Sub

Private Sub txtrazao_LostFocus(Index As Integer)
letraMaiusculo (txtrazao(0).Text)
End Sub

Private Sub tXTsITE_Change(Index As Integer)
AddoutRecibo
End Sub

Private Sub tXTtELEFONE_Change(Index As Integer)
AddoutRecibo
End Sub

Private Sub tXTtELEFONErecibo_Change(Index As Integer)
AddoutRecibo
End Sub

Private Sub TXTuF_Change(Index As Integer)
AddoutRecibo
End Sub

Public Function AddoutRecibo()
Txtendercorecibo(1).Text = TxtEndereco(0).Text
numerorecibo(1).Text = txtN(0).Text
textComplementoRecibo(1).Text = TxtComplemento(0).Text
tXTbAIRROrecibo(1).Text = tXTbAIRRO(0).Text
TXTcIDADErecibo(1).Text = TXTcIDADE(0).Text
tXTcEPrecibo(1).Text = tXTcEP(0).Text
TXTuFrecibo(1).Text = TXTuF(0).Text
TxtreciboCnpjrecibo.Text = TXTcPF.Text
tXTtELEFONErecibo(1).Text = tXTtELEFONE(0).Text
tXTcONTATOrecibo(1).Text = tXTcONTATO(0).Text
tXTfAXrecibo(1).Text = tXTfAX(0).Text
tXTcELULARrecibo(1).Text = tXTcELULAR(0).Text
tXTsITErecibo(1).Text = tXTsITE(0).Text
tEXTeMAILrecibo(1).Text = txtemail(0).Text

End Function

Public Sub LetrasMaiuscolas()
txtrazao(0).Text = Format(txtrazao(0).Text, ">")
textfantazia(0).Text = Format(textfantazia(0).Text, ">")

Txtendercorecibo(1).Text = Format(Txtendercorecibo(1).Text, ">")
TxtEndereco(0).Text = Format(TxtEndereco(0).Text, ">")
numerorecibo(1).Text = Format(numerorecibo(1).Text, ">")
txtN(0).Text = Format(txtN(0).Text, ">")
textComplementoRecibo(1).Text = Format(textComplementoRecibo(1).Text, ">")
TxtComplemento(0).Text = Format(TxtComplemento(0).Text, ">")
tXTbAIRROrecibo(1).Text = Format(tXTbAIRROrecibo(1).Text, ">")
tXTbAIRRO(0).Text = Format(tXTbAIRRO(0).Text, ">")
TXTcIDADErecibo(1).Text = Format(TXTcIDADErecibo(1).Text, ">")
TXTcIDADE(0).Text = Format(TXTcIDADE(0).Text, ">")
tXTcEPrecibo(1).Text = Format(tXTcEPrecibo(1).Text, ">")
tXTcEP(0).Text = Format(tXTcEP(0).Text, ">")
TXTuFrecibo(1).Text = Format(TXTuFrecibo(1).Text, ">")
TXTuF(0).Text = Format(TXTuF(0).Text, ">")
TxtreciboCnpjrecibo.Text = Format(TxtreciboCnpjrecibo.Text, ">")
TXTcPF.Text = Format(TXTcPF.Text, ">")
tXTtELEFONErecibo(1).Text = Format(tXTtELEFONErecibo(1).Text, ">")
tXTtELEFONE(0).Text = Format(tXTtELEFONE(0).Text, ">")
tXTcONTATOrecibo(1).Text = Format(tXTcONTATOrecibo(1).Text, ">")
tXTcONTATO(0).Text = Format(tXTcONTATO(0).Text, ">")
tXTfAXrecibo(1).Text = Format(tXTfAXrecibo(1).Text, ">")
tXTfAX(0).Text = Format(tXTfAX(0).Text, ">")
tXTcELULARrecibo(1).Text = Format(tXTcELULARrecibo(1).Text, ">")
tXTcELULAR(0).Text = Format(tXTcELULAR(0).Text, ">")
tXTsITErecibo(1).Text = Format(tXTsITErecibo(1).Text, ">")
tXTsITE(0).Text = Format(tXTsITE(0).Text, ">")
tEXTeMAILrecibo(1).Text = Format(tEXTeMAILrecibo(1).Text, ">")
txtemail(0).Text = Format(txtemail(0).Text, ">")
End Sub



Public Sub minuscolas()
Txtendercorecibo(1).Text = Format(Txtendercorecibo(1).Text, "<")
TxtEndereco(0).Text = Format(TxtEndereco(0).Text, "<")
numerorecibo(1).Text = Format(numerorecibo(1).Text, "<")
txtN(0).Text = Format(txtN(0).Text, "<")
textComplementoRecibo(1).Text = Format(textComplementoRecibo(1).Text, "<")
TxtComplemento(0).Text = Format(TxtComplemento(0).Text, "<")
tXTbAIRROrecibo(1).Text = Format(tXTbAIRROrecibo(1).Text, "<")
tXTbAIRRO(0).Text = Format(tXTbAIRRO(0).Text, "<")
TXTcIDADErecibo(1).Text = Format(TXTcIDADErecibo(1).Text, "<")
TXTcIDADE(0).Text = Format(TXTcIDADE(0).Text, "<")
tXTcEPrecibo(1).Text = Format(tXTcEPrecibo(1).Text, "<")
tXTcEP(0).Text = Format(tXTcEP(0).Text, "<")
TXTuFrecibo(1).Text = Format(TXTuFrecibo(1).Text, "<")
TXTuF(0).Text = Format(TXTuF(0).Text, "<")
TxtreciboCnpjrecibo.Text = Format(TxtreciboCnpjrecibo.Text, "<")
TXTcPF.Text = Format(TXTcPF.Text, "<")
tXTtELEFONErecibo(1).Text = Format(tXTtELEFONErecibo(1).Text, "<")
tXTtELEFONE(0).Text = Format(tXTtELEFONE(0).Text, "<")
tXTcONTATOrecibo(1).Text = Format(tXTcONTATOrecibo(1).Text, "<")
tXTcONTATO(0).Text = Format(tXTcONTATO(0).Text, "<")
tXTfAXrecibo(1).Text = Format(tXTfAXrecibo(1).Text, "<")
tXTfAX(0).Text = Format(tXTfAX(0).Text, "<")
tXTcELULARrecibo(1).Text = Format(tXTcELULARrecibo(1).Text, "<")
tXTcELULAR(0).Text = Format(tXTcELULAR(0).Text, "<")
tXTsITErecibo(1).Text = Format(tXTsITErecibo(1).Text, "<")
tXTsITE(0).Text = Format(tXTsITE(0).Text, "<")
tEXTeMAILrecibo(1).Text = Format(tEXTeMAILrecibo(1).Text, "<")
txtemail(0).Text = Format(txtemail(0).Text, "<")
End Sub

Public Sub controleLetrass()
If CmbFormatacao = "LETRAS MAIUSCÚLAS" Then
LetrasMaiuscolas
ElseIf CmbFormatacao.Text = "minuscúlas" Then
minuscolas
Else
End If

End Sub
