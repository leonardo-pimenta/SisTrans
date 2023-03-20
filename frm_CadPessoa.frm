VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_CadPessoa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SisTrans - Cadastro de Pessoas"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10965
   ControlBox      =   0   'False
   Icon            =   "frm_CadPessoa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btn_Excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   2160
      Picture         =   "frm_CadPessoa.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Excluir"
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton btn_Cartao 
      Caption         =   "&Cartão"
      Height          =   855
      Left            =   6480
      Picture         =   "frm_CadPessoa.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Editar"
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton btn_Veiculo 
      Caption         =   "Veí&culo"
      Height          =   855
      Left            =   7440
      Picture         =   "frm_CadPessoa.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Editar"
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "E&ditar"
      Height          =   855
      Left            =   1200
      Picture         =   "frm_CadPessoa.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Editar"
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   9840
      Picture         =   "frm_CadPessoa.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Sal&var"
      Height          =   855
      Left            =   240
      Picture         =   "frm_CadPessoa.frx":1854
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Salvar"
      Top             =   5760
      Width           =   855
   End
   Begin VB.Frame fra_Cadastro 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   120
      TabIndex        =   40
      Top             =   120
      Width           =   10695
      Begin VB.ComboBox COMBO_ESPECIFICACAO 
         Height          =   315
         ItemData        =   "frm_CadPessoa.frx":1C96
         Left            =   1560
         List            =   "frm_CadPessoa.frx":1CA3
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TXT_NOME 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   5400
         MaxLength       =   50
         TabIndex        =   3
         Top             =   360
         Width           =   5175
      End
      Begin VB.TextBox TXT_NOME_GUERRA 
         Height          =   300
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin MSMask.MaskEdBox TXT_DT_DESEMBARQUE 
         Height          =   300
         Left            =   8520
         TabIndex        =   23
         Top             =   2880
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   -2147483644
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TXT_DT_EMBARQUE 
         Height          =   300
         Left            =   6600
         TabIndex        =   22
         Top             =   2880
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   -2147483644
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TXT_DT_PROMOCAO 
         Height          =   300
         Left            =   4560
         TabIndex        =   21
         Top             =   2880
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   -2147483644
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TXT_DT_NASCIMENTO 
         Height          =   300
         Left            =   2160
         TabIndex        =   20
         Top             =   2880
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   -2147483644
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TXT_NIP 
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   10
         Mask            =   "99.9999.99"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CMB_UF_CNH 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frm_CadPessoa.frx":1CC9
         Left            =   9960
         List            =   "frm_CadPessoa.frx":1D21
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2040
         Width           =   615
      End
      Begin MSMask.MaskEdBox Txt_Val_CNH 
         Height          =   300
         Left            =   8760
         TabIndex        =   17
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   -2147483644
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TXT_DT_VAL_ID 
         Height          =   300
         Left            =   4080
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox COMBO_CORPO_QUADRO_ESPECIALIDADE 
         Height          =   315
         Left            =   6000
         TabIndex        =   7
         Text            =   "COMBO_CORPO_QUADRO_ESPECIALIDADE"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton Exc_subespecialidade 
         Caption         =   "->"
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
         Height          =   225
         Left            =   9840
         Picture         =   "frm_CadPessoa.frx":1D94
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1245
         Width           =   375
      End
      Begin VB.CommandButton Inc_Subespecialidade 
         Caption         =   "<-"
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
         Height          =   255
         Left            =   9840
         Picture         =   "frm_CadPessoa.frx":21D6
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1005
         Width           =   375
      End
      Begin MSDataListLib.DataCombo DBCOM_COD_POSTGRAD_POSTGRADCATFUNC 
         Height          =   315
         Left            =   3600
         TabIndex        =   6
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Post_Grad_CatFunc"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DBCOM_COD_OM 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   1200
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "SIGLA"
         Text            =   ""
      End
      Begin VB.TextBox TXT_COMPLEMENTO 
         Height          =   285
         Left            =   600
         MaxLength       =   50
         TabIndex        =   39
         Top             =   5640
         Width           =   10095
      End
      Begin VB.Frame Frame6 
         Caption         =   "Endereço:"
         Height          =   2295
         Left            =   0
         TabIndex        =   61
         Top             =   3240
         Width           =   10575
         Begin VB.ComboBox CMB_UF_ENDERECO 
            Height          =   315
            ItemData        =   "frm_CadPessoa.frx":2618
            Left            =   8160
            List            =   "frm_CadPessoa.frx":2670
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1080
            Width           =   615
         End
         Begin MSMask.MaskEdBox TXT_TELEFONE2 
            Height          =   300
            Left            =   1440
            TabIndex        =   30
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   13
            Mask            =   "(99)9999-9999"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TXT_TELEFONE1 
            Height          =   300
            Left            =   120
            TabIndex        =   29
            Top             =   1800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   13
            Mask            =   "(99)9999-9999"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TXT_CEP 
            Height          =   300
            Left            =   9000
            TabIndex        =   28
            Top             =   1080
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Format          =   "  .   -"
            Mask            =   "99.999-999"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TXT_E_MAIL2 
            Height          =   285
            Left            =   8160
            MaxLength       =   50
            TabIndex        =   32
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox TXT_CIDADE 
            Height          =   285
            Left            =   3960
            MaxLength       =   50
            TabIndex        =   26
            Top             =   1080
            Width           =   3735
         End
         Begin VB.TextBox TXT_BAIRRO 
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   25
            Top             =   1080
            Width           =   3735
         End
         Begin VB.TextBox TXT_LOGRADOURO 
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   24
            Top             =   480
            Width           =   7215
         End
         Begin VB.TextBox TXT_E_MAIL1 
            Height          =   285
            Left            =   5040
            MaxLength       =   50
            TabIndex        =   31
            Top             =   1680
            Width           =   3015
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Telefone2:"
            Height          =   195
            Left            =   1440
            TabIndex        =   70
            Top             =   1440
            Width           =   765
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
            Height          =   195
            Left            =   9000
            TabIndex        =   69
            Top             =   840
            Width           =   360
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Left            =   3960
            TabIndex        =   68
            Top             =   840
            Width           =   540
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   840
            Width           =   450
         End
         Begin VB.Label Label25 
            Caption         =   "UF:"
            Height          =   195
            Left            =   8160
            TabIndex        =   66
            Top             =   840
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Telefone1:"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   1440
            Width           =   765
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro (rua, avenida, alameda, etc.) :"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   2970
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "E-mail1:"
            Height          =   195
            Left            =   5040
            TabIndex        =   63
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "E-mail2:"
            Height          =   195
            Left            =   8160
            TabIndex        =   62
            Top             =   1440
            Width           =   555
         End
      End
      Begin VB.Frame Frame5 
         Height          =   135
         Left            =   0
         TabIndex        =   60
         Top             =   2400
         Width           =   10575
      End
      Begin VB.Frame Frame4 
         Height          =   135
         Left            =   0
         TabIndex        =   59
         Top             =   1560
         Width           =   10575
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Left            =   0
         TabIndex        =   58
         Top             =   720
         Width           =   10575
      End
      Begin VB.Frame Frame3 
         Height          =   135
         Left            =   120
         TabIndex        =   43
         Top             =   5880
         Width           =   10575
      End
      Begin VB.TextBox TXT_NUM_CNH 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         MaxLength       =   50
         TabIndex        =   15
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox TXT_NUMREG_CNH 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6960
         MaxLength       =   50
         TabIndex        =   16
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox TXT_IDENTIDADE 
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   12
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox TXT_ORGAOEXPED_ID 
         Height          =   285
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   13
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox TXT_CODINOME 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox TXT_PASEP 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   11
         TabIndex        =   19
         Top             =   2880
         Width           =   1575
      End
      Begin MSMask.MaskEdBox TXT_CPF 
         Height          =   300
         Left            =   0
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "999.999.999-99"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo DBCOM_SUBESPECIALIDADE 
         Height          =   315
         Left            =   8520
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         BackColor       =   -2147483644
         ListField       =   "SIGLA_SUBESPECIALIDADE"
         BoundColumn     =   "SIGLA_SUBESPECIALIDADE"
         Text            =   ""
      End
      Begin VB.Label Label31 
         Caption         =   "Subespecialidade"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   8520
         TabIndex        =   81
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "CPF:"
         Height          =   195
         Left            =   0
         TabIndex        =   80
         Top             =   120
         Width           =   345
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   5400
         TabIndex        =   79
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label19 
         Caption         =   "Especificação:"
         Height          =   195
         Left            =   1560
         TabIndex        =   78
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome de guerra:"
         Height          =   195
         Left            =   3360
         TabIndex        =   77
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label Label18 
         Caption         =   "Espec."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   7440
         TabIndex        =   76
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Quadro"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6720
         TabIndex        =   75
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Corpo"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6000
         TabIndex        =   74
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Validade:"
         Height          =   195
         Left            =   8760
         TabIndex        =   73
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Obs:"
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   5640
         Width           =   330
      End
      Begin VB.Label Label27 
         Caption         =   "Validade:"
         Height          =   195
         Left            =   4080
         TabIndex        =   71
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label Label9 
         Caption         =   "Data de embarque:"
         Height          =   195
         Left            =   6600
         TabIndex        =   57
         Top             =   2640
         Width           =   1515
      End
      Begin VB.Label Label10 
         Caption         =   "Data do desembarque:"
         Height          =   195
         Left            =   8520
         TabIndex        =   56
         Top             =   2640
         Width           =   1620
      End
      Begin VB.Label Label3 
         Caption         =   "Posto/Graduação:"
         Height          =   255
         Left            =   3600
         TabIndex        =   55
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Data promoção:"
         Height          =   195
         Left            =   4560
         TabIndex        =   54
         Top             =   2640
         Width           =   1155
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         Height          =   195
         Left            =   9960
         TabIndex        =   53
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label28 
         Caption         =   "Nº Cart. Habilitação:"
         Height          =   195
         Left            =   5280
         TabIndex        =   52
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label Label26 
         Caption         =   "Nº Reg. Habilitação:"
         Height          =   195
         Left            =   6960
         TabIndex        =   51
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label24 
         Caption         =   "Identidade:"
         Height          =   195
         Left            =   1320
         TabIndex        =   50
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label Label23 
         Caption         =   "Org.Exp:"
         Height          =   195
         Left            =   2760
         TabIndex        =   49
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label Label22 
         Caption         =   "Data Nascimento:"
         Height          =   195
         Left            =   2160
         TabIndex        =   48
         Top             =   2640
         Width           =   1275
      End
      Begin VB.Label Label21 
         Caption         =   "Codnome:"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "NIP:"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   960
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "OM:"
         Height          =   195
         Left            =   1320
         TabIndex        =   45
         Top             =   960
         Width           =   300
      End
      Begin VB.Label Label15 
         Caption         =   "PASEP:"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   2640
         Width           =   555
      End
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Complemento:"
      Height          =   195
      Left            =   240
      TabIndex        =   42
      Top             =   4440
      Width           =   1005
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Bairro:"
      Height          =   195
      Left            =   240
      TabIndex        =   41
      Top             =   4800
      Width           =   450
   End
End
Attribute VB_Name = "frm_CadPessoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub btn_Cartao_Click()
'Envia os dados do form atual para o de Cartao p/
'ser feitas as deidas operacoes
vgl_TipoResponsavel = "CPF"
vgl_PostoGrad = DBCOM_COD_POSTGRAD_POSTGRADCATFUNC.Text
var_SQL = "SELECT * FROM tab_trans_aux_posto_Tarja "
var_SQL = var_SQL + "WHERE PostoGrad = '" & vgl_PostoGrad & "'"
Set DS = New Recordset
DS.Open var_SQL, cnConexao
If DS.RecordCount = 0 Or DS!tarjacartao = "" Then
   var_MSG = "A graduação a que pertence o militar não possui permissão de se "
   var_MSG = var_MSG + "emitir um cartão de estacionamento ou a graduação a que "
   var_MSG = var_MSG + "pertence não foi cadastrada junto a uma tarja."
   MsgBox var_MSG, vbOKOnly + vbInformation, "SisTrans"
   Exit Sub
Else
   frm_Cartao_Emitir.Txt_Tarja.Text = DS!tarjacartao
   frm_Cartao_Emitir.txt_Codigo = TXT_CPF.Text
   frm_Cartao_Emitir.txt_Descricao = TXT_NOME.Text
End If
End Sub
Private Sub btn_Excluir_Click()
On Error GoTo Error
If MsgBox("Deseja excluir o registro selecionado?", vbYesNo + vbQuestion, "SisTrans") = vbYes Then
    cmdEditar.Enabled = False
    btn_Excluir.Enabled = False
    cmdSalvar.Enabled = False
    SQL = "select * FROM TAB_GER_PESSOA"
    SQL = SQL + " WHERE CPF='" & TXT_CPF & "'"
    Set DS = New Recordset
    DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
    SQL = "select * FROM TAB_GER_PESSOA_SUBESPECIALIDADE"
    SQL = SQL + " WHERE CPF_PESSOA='" & TXT_CPF & "'"
    Set ds1 = New Recordset
    ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
    Do While Not ds1.EOF
       ds1.Delete
       ds1.MoveNext
    Loop
    DS.Delete
End If
Call LIMPAR
TXT_CPF.Mask = "              "
TXT_CPF.Mask = "999.999.999-99"

MsgBox "Registro deletado.", vbInformation, "SISTRANS"
cmdSalvar.Enabled = True
Exit Sub
Error:
MsgBox "Não é possível excluir o registro, ele pode fazer parte de um ou mais relacionamentos.Ex.: multa x veículo, veículo x proprietário. Caso seja realmente necessária a exclusão, contate o administrador do Banco de Dados.", vbOKOnly + vbInformation, "SisTrans"
cmdSair.SetFocus
End Sub
Private Sub btn_Limpar_Click()
TXT_CPF = "___.___.___-__"
VarAuxiliar = "Adicionando"
SQL = "select * FROM TAB_GER_PESSOA WHERE Auxiliar='" & VarAuxiliar & "'"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount <> 0 Then
   Do While Not DS.EOF
      DS.Delete
      DS.MoveNext
   Loop
End If
Call LIMPAR
TXT_CPF.Enabled = True
cmdSalvar.Enabled = False
btn_Excluir.Enabled = False
btn_Cartao.Enabled = False
btn_Veiculo.Enabled = False
cmdEditar.Enabled = False
'btn_Limpar.Enabled = False
fra_Cadastro.Enabled = True

TXT_CPF.Mask = "              "
TXT_CPF.Mask = "999.999.999-99"
TXT_CPF.SetFocus
End Sub
Private Sub btn_Veiculo_Click()
frm_CadVeiculo.vgl_Responsavel = 1
frm_CadVeiculo.opt_Pessoa.Value = True
frm_CadVeiculo.txt_Responsavel = TXT_CPF.Text
frm_CadVeiculo.txt_Descricao = TXT_NOME.Text

frm_CadVeiculo.Show

End Sub
Private Sub cmdAdicionar_Click()
Call LIMPAR
End Sub

Private Sub cmdEditar_Click()
TXT_CPF.Enabled = False
cmdSalvar.Enabled = True
'btn_Limpar.Enabled = True

cmdEditar.Enabled = False
btn_Excluir.Enabled = False
btn_Cartao.Enabled = False
btn_Veiculo.Enabled = False
fra_Cadastro.Enabled = True
COMBO_ESPECIFICACAO.SetFocus

End Sub
Private Sub cmdSair_Click()
VarAuxiliar = "Adicionando"
SQL = "select * FROM TAB_GER_PESSOA WHERE Auxiliar='" & VarAuxiliar & "'"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount <> 0 Then
   Do While Not DS.EOF
      DS.Delete
      DS.MoveNext
   Loop
End If
Call LIMPAR
cmdSalvar.Enabled = True
Unload Me
End Sub
Private Sub cmdSalvar_Click()
If TXT_CPF.Text = "___.___.___-__" Then
   MsgBox "Digite o CPF.", vbInformation, "SISTRANS"
   Exit Sub
End If
If DBCOM_COD_POSTGRAD_POSTGRADCATFUNC.Text = "" Then
   MsgBox "Selecione o Posto/Graduação.", vbInformation, "SISTRANS"
   Exit Sub
End If
If COMBO_CORPO_QUADRO_ESPECIALIDADE.Text = "" Then
   MsgBox "Selecione o Corpo/Quadro/Especialidade.", vbInformation, "SISTRANS"
   Exit Sub
End If
If DBCOM_COD_OM.Text = "" Then
   MsgBox "Selecione o Código da OM.", vbInformation, "SISTRANS"
   Exit Sub
End If
If VarAcao = "Registro Novo" Then
   Call Salvar
   MsgBox "Registro salvo.", vbInformation, "SISTRANS"
Else
   Call Salvar
   MsgBox "Registro alterado.", vbInformation, "SISTRANS"
End If
Call Form_Load
End Sub
Private Sub Salvar()

SQL = "select * FROM TAB_GER_PESSOA"
SQL = SQL + " WHERE CPF='" & TXT_CPF & "'"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

If DS.RecordCount = 0 Then
    DS.AddNew
    DS!CPF = TXT_CPF.Text
End If

DS!identidade = TXT_IDENTIDADE
DS!OrgaoExped_Id = TXT_ORGAOEXPED_ID
DS!Dt_Val_Id = TXT_DT_VAL_ID
DS!ESPECIFICACAO = COMBO_ESPECIFICACAO.Text
DS!NIP = TXT_NIP
DS!PASEP = ""
DS!DT_EMBARQUE = "  /  /    "
DS!DT_DESEMBARQUE = "  /  /    "
DS!DT_NASCIMENTO = "  /  /    "
DS!DT_PROMOÇÃO = "  /  /    "
DS!NOME = TXT_NOME
DS!NOME_GUERRA = TXT_NOME_GUERRA
DS!CODINOME = TXT_CODINOME
DS!NUM_CNH = TXT_NUM_CNH
DS!VAL_CNH = Txt_Val_CNH
DS!UF_CNH = CMB_UF_CNH.Text
DS!UF_ENDERECO = CMB_UF_ENDERECO.Text
DS!NUMREG_CNH = TXT_NUMREG_CNH
DS!TELEFONE1 = TXT_TELEFONE1
DS!TELEFONE2 = TXT_TELEFONE2
DS!E_MAIL1 = TXT_E_MAIL1
DS!E_MAIL2 = TXT_E_MAIL2
DS!LOGRADOURO = TXT_LOGRADOURO
DS!BAIRRO = TXT_BAIRRO
DS!CIDADE = TXT_CIDADE
DS!CEP = TXT_CEP
DS!COMPLEMENTO = TXT_COMPLEMENTO
If COMBO_CORPO_QUADRO_ESPECIALIDADE.Text <> "" Then
   DS!COD_CORPO_QUADRO_ESPECIALIDADE = COMBO_CORPO_QUADRO_ESPECIALIDADE.ItemData(COMBO_CORPO_QUADRO_ESPECIALIDADE.ListIndex)
End If
If DBCOM_COD_POSTGRAD_POSTGRADCATFUNC <> "" Then
   DS!cod_postGrad_PostGradCatFunc = DBCOM_COD_POSTGRAD_POSTGRADCATFUNC
End If

DS!AUXILIAR = "Editando"

SQL = "select * FROM TAB_GER_OM"
SQL = SQL + " WHERE SIGLA='" & DBCOM_COD_OM & "'"
Set ds1 = New Recordset
ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
If ds1.RecordCount <> 0 Then
   If ds1(0) <> "" Then
      DS!COD_OM = ds1(0)
   End If
End If

DS.UpdateBatch adAffectAll

Call LIMPAR
TXT_CPF.Mask = "              "
TXT_CPF.Mask = "999.999.999-99"

End Sub
Private Sub LIMPAR()

TXT_IDENTIDADE = ""
TXT_ORGAOEXPED_ID = ""
TXT_DT_VAL_ID.Mask = "          "
TXT_DT_VAL_ID.Mask = "99/99/9999"
COMBO_ESPECIFICACAO = ""
TXT_NIP.Mask = "          "
TXT_NIP.Mask = "99.9999.99"
TXT_PASEP = ""
TXT_DT_EMBARQUE.Mask = "          "
TXT_DT_EMBARQUE.Mask = "99/99/9999"
TXT_DT_DESEMBARQUE.Mask = "          "
TXT_DT_DESEMBARQUE.Mask = "99/99/9999"
TXT_DT_NASCIMENTO.Mask = "          "
TXT_DT_NASCIMENTO.Mask = "99/99/9999"
TXT_NOME = ""
TXT_NOME_GUERRA = ""
TXT_CODINOME = ""
TXT_NUM_CNH = ""
Txt_Val_CNH.Mask = "          "
Txt_Val_CNH.Mask = "99/99/9999"
CMB_UF_CNH.ListIndex = 0
CMB_UF_ENDERECO.ListIndex = 0
TXT_NUMREG_CNH = ""
TXT_TELEFONE1.Mask = "             "
TXT_TELEFONE1.Mask = "(99)9999-9999"
TXT_TELEFONE2.Mask = "             "
TXT_TELEFONE2.Mask = "(99)9999-9999"
TXT_E_MAIL1 = ""
TXT_E_MAIL2 = ""
TXT_LOGRADOURO = ""
TXT_BAIRRO = ""
TXT_CIDADE = ""
TXT_CEP.Mask = "          "
TXT_CEP.Mask = "99.999-999"
TXT_COMPLEMENTO = ""
TXT_DT_PROMOCAO.Mask = "          "
TXT_DT_PROMOCAO.Mask = "99/99/9999"
DBCOM_COD_OM = ""
DBCOM_SUBESPECIALIDADE = ""
DBCOM_COD_POSTGRAD_POSTGRADCATFUNC = ""
COMBO_CORPO_QUADRO_ESPECIALIDADE.ListIndex = -1
End Sub
Private Sub DGRID_CORPO_QUADRO_ESPECIALIDADE_Click()
VarEspecialidade = DGRID_CORPO_QUADRO_ESPECIALIDADE.Columns(3)
End Sub

Private Sub Exc_subespecialidade_Click()
SQL = "select SIGLA_SUBESPECIALIDADE "
SQL = SQL + " FROM TAB_GER_PESSOA_SUBESPECIALIDADE"
SQL = SQL + " where CPF_PESSOA='" & TXT_CPF & "'"
SQL = SQL + " AND SIGLA_SUBESPECIALIDADE='" & DBCOM_SUBESPECIALIDADE.Text & "'"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount <> 0 Then
   DS.Delete
End If
'Combo SUBESPECIALIDADE
DBCOM_SUBESPECIALIDADE = ""
SQL = "select SIGLA_SUBESPECIALIDADE "
SQL = SQL + " FROM TAB_GER_PESSOA_SUBESPECIALIDADE"
SQL = SQL + " where CPF_PESSOA='" & TXT_CPF & "'"
SQL = SQL + " order by SIGLA_SUBESPECIALIDADE"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set DBCOM_SUBESPECIALIDADE.RowSource = DS
End Sub
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Call LIMPAR
TXT_CPF.Mask = "              "
TXT_CPF.Mask = "999.999.999-99"
CMB_UF_CNH.ListIndex = 0
CMB_UF_ENDERECO.ListIndex = 0
TXT_CPF.Enabled = True
'Combo OM
SQL = "select SIGLA FROM TAB_GER_OM"
SQL = SQL + " order by SIGLA"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set DBCOM_COD_OM.RowSource = DS

'Combo Posto/Graduação
SQL = "select Post_Grad_CatFunc FROM tab_ger_post_grad_catfunc"
SQL = SQL + " order by HIERARQUIA"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set DBCOM_COD_POSTGRAD_POSTGRADCATFUNC.RowSource = DS


'COMBO CORPO-QUADRO-ESPECIALIDADE
SQL = "SELECT TAB_GER_CORPO.SIGLA,TAB_GER_QUADRO.SIGLA,TAB_GER_ESPECIALIDADE.SIGLA,"
SQL = SQL + " TAB_GER_CORPO_QUADRO_ESPECIALIDADE.COD"
SQL = SQL + " FROM  TAB_GER_CORPO_QUADRO_ESPECIALIDADE,TAB_GER_CORPO,TAB_GER_QUADRO,TAB_GER_ESPECIALIDADE"
SQL = SQL + " Where TAB_GER_CORPO_QUADRO_ESPECIALIDADE.COD_ESPECIALIDADE = TAB_GER_ESPECIALIDADE.COD"
SQL = SQL + " AND TAB_GER_CORPO_QUADRO_ESPECIALIDADE.COD_QUADRO=TAB_GER_QUADRO.COD"
SQL = SQL + " AND TAB_GER_CORPO_QUADRO_ESPECIALIDADE.COD_Corpo=TAB_GER_CORPO.COD"
SQL = SQL + " ORDER BY TAB_GER_CORPO.SIGLA,TAB_GER_QUADRO.SIGLA,TAB_GER_ESPECIALIDADE.SIGLA"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

COMBO_CORPO_QUADRO_ESPECIALIDADE.Clear
Do While Not DS.EOF
   COMBO_CORPO_QUADRO_ESPECIALIDADE.AddItem ((DS(0)) + " - " + DS(1) + " - " + DS(2))
   COMBO_CORPO_QUADRO_ESPECIALIDADE.ItemData(COMBO_CORPO_QUADRO_ESPECIALIDADE.NewIndex) = DS(3)
   DS.MoveNext
Loop
DS.Close
cmdSalvar.Enabled = False
btn_Excluir.Enabled = False
btn_Cartao.Enabled = False
btn_Veiculo.Enabled = False
cmdEditar.Enabled = False
'btn_Limpar.Enabled = False

End Sub
Private Sub Form_Unload(Cancel As Integer)
Call cmdSair_Click
End Sub

Private Sub Inc_subespecialidade_Click()
frm_Subespecialidade.Show
End Sub
Private Sub TXT_BAIRRO_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_CIDADE_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_CODINOME_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_COMPLEMENTO_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_CPF_GotFocus()
Call LIMPAR
TXT_CPF = "___.___.___-__"
VarAuxiliar = "Adicionando"
SQL = "select * FROM TAB_GER_PESSOA WHERE Auxiliar='" & VarAuxiliar & "'"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount <> 0 Then
   Do While Not DS.EOF
      DS.Delete
      DS.MoveNext
   Loop
End If
End Sub

Private Sub txt_CPF_LostFocus()

If TXT_CPF <> "___.___.___-__" Then
   
    If Len(TXT_CPF) < 14 Then
    MsgBox "CPF inválido!", vbInformation
    TXT_CPF.SetFocus
    Exit Sub
    End If
    

   SQL = "select * FROM TAB_GER_PESSOA"
   SQL = SQL + " where CPF='" & TXT_CPF & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   If DS.RecordCount = 1 Then
   
      VarAcao = "Registro Velho"
       
      TXT_CPF = DS!CPF
      TXT_IDENTIDADE = DS!identidade
      TXT_ORGAOEXPED_ID = DS!OrgaoExped_Id
      TXT_DT_VAL_ID = DS!Dt_Val_Id
      COMBO_ESPECIFICACAO.Text = DS!ESPECIFICACAO
      TXT_NIP = DS!NIP
      TXT_PASEP = DS!PASEP
      TXT_DT_EMBARQUE = DS!DT_EMBARQUE
      TXT_DT_DESEMBARQUE = DS!DT_DESEMBARQUE
      TXT_DT_NASCIMENTO = DS!DT_NASCIMENTO
      TXT_NOME = DS!NOME
      TXT_NOME_GUERRA = DS!NOME_GUERRA
      TXT_CODINOME = DS!CODINOME
      TXT_NUM_CNH = DS!NUM_CNH
      Txt_Val_CNH = DS!VAL_CNH
      TXT_NUMREG_CNH = DS!NUMREG_CNH
      TXT_TELEFONE1 = DS!TELEFONE1
      TXT_TELEFONE2 = DS!TELEFONE2
      TXT_E_MAIL1 = DS!E_MAIL1
      TXT_E_MAIL2 = DS!E_MAIL2
      TXT_LOGRADOURO = DS!LOGRADOURO
      TXT_BAIRRO = DS!BAIRRO
      TXT_CIDADE = DS!CIDADE
      TXT_CEP = DS!CEP
      TXT_COMPLEMENTO = DS!COMPLEMENTO
      'Posiciona a combo UF no item do arquivo.
   
      VarIndex = 0
      CMB_UF_CNH.ListIndex = 0
      Do While Not (CMB_UF_CNH.Text = DS(15))
         If VarIndex <= (CMB_UF_CNH.ListCount - 1) Then
            VarIndex = (CMB_UF_CNH.ListIndex + 1)
            CMB_UF_CNH.ListIndex = VarIndex
         End If
      Loop
      CMB_UF_CNH.ListIndex = VarIndex
      
      'Posiciona a combo UF no item do arquivo.
   
      VarIndex = 0
      CMB_UF_ENDERECO.ListIndex = 0
      Do While Not (CMB_UF_ENDERECO.Text = DS(16))
         If VarIndex <= (CMB_UF_ENDERECO.ListCount - 1) Then
            VarIndex = (CMB_UF_ENDERECO.ListIndex + 1)
            CMB_UF_ENDERECO.ListIndex = VarIndex
         End If
      Loop
      CMB_UF_ENDERECO.ListIndex = VarIndex
      
      If DS!cod_postGrad_PostGradCatFunc <> "" Then
         DBCOM_COD_POSTGRAD_POSTGRADCATFUNC = DS!cod_postGrad_PostGradCatFunc
      End If
      TXT_DT_PROMOCAO = DS!DT_PROMOÇÃO
   
      SQL = "select SIGLA_SUBESPECIALIDADE FROM TAB_GER_PESSOA_SUBESPECIALIDADE"
      SQL = SQL + " WHERE CPF_PESSOA='" & TXT_CPF & "'"
      Set ds1 = New Recordset
      ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      Set DBCOM_SUBESPECIALIDADE.RowSource = ds1
      
      SQL = "select * FROM TAB_GER_OM,TAB_GER_PESSOA"
      SQL = SQL + " WHERE TAB_GER_PESSOA.COD_OM=TAB_GER_OM.COD"
      SQL = SQL + " AND CPF='" & TXT_CPF & "'"
      Set ds1 = New Recordset
      ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      If ds1.RecordCount <> 0 Then
         DBCOM_COD_OM = ds1(1)
      End If
               
      'COMBO CORPO-QUADRO-ESPECIALIDADE
      SQL = "SELECT TAB_GER_CORPO.SIGLA,TAB_GER_QUADRO.SIGLA,TAB_GER_ESPECIALIDADE.SIGLA,"
      SQL = SQL + " TAB_GER_CORPO_QUADRO_ESPECIALIDADE.COD"
      SQL = SQL + " FROM  TAB_GER_CORPO_QUADRO_ESPECIALIDADE,TAB_GER_CORPO,TAB_GER_QUADRO,"
      SQL = SQL + " TAB_GER_ESPECIALIDADE,TAB_GER_PESSOA"
      SQL = SQL + " Where TAB_GER_CORPO_QUADRO_ESPECIALIDADE.COD_ESPECIALIDADE = TAB_GER_ESPECIALIDADE.COD"
      SQL = SQL + " AND TAB_GER_CORPO_QUADRO_ESPECIALIDADE.COD_QUADRO=TAB_GER_QUADRO.COD"
      SQL = SQL + " AND TAB_GER_CORPO_QUADRO_ESPECIALIDADE.COD_Corpo=TAB_GER_CORPO.COD"
      SQL = SQL + " AND TAB_GER_PESSOA.COD_CORPO_QUADRO_ESPECIALIDADE=TAB_GER_CORPO_QUADRO_ESPECIALIDADE.COD "
      SQL = SQL + " AND CPF='" & TXT_CPF & "'"
      SQL = SQL + " ORDER BY TAB_GER_CORPO.SIGLA,TAB_GER_QUADRO.SIGLA,TAB_GER_ESPECIALIDADE.SIGLA"
      Set DS = New Recordset
      DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      If DS.RecordCount <> 0 Then
         'Posiciona a combo QUADRO no item do arquivo.

         VarIndex = 0
         COMBO_CORPO_QUADRO_ESPECIALIDADE.ListIndex = 0
         Do While Not (COMBO_CORPO_QUADRO_ESPECIALIDADE.ItemData(COMBO_CORPO_QUADRO_ESPECIALIDADE.ListIndex)) = DS(3)
            If VarIndex <= (COMBO_CORPO_QUADRO_ESPECIALIDADE.ListCount - 1) Then
               VarIndex = (COMBO_CORPO_QUADRO_ESPECIALIDADE.ListIndex + 1)
               COMBO_CORPO_QUADRO_ESPECIALIDADE.ListIndex = VarIndex
            End If
         Loop
         COMBO_CORPO_QUADRO_ESPECIALIDADE.ListIndex = VarIndex
      End If
      'btn_Limpar.Enabled = True
      cmdSalvar.Enabled = False
      cmdEditar.Enabled = True
      btn_Excluir.Enabled = True
      fra_Cadastro.Enabled = False
      btn_Cartao.Enabled = True
      btn_Veiculo.Enabled = True
      
    Else
      SQL = "select * FROM TAB_GER_PESSOA"
      Set DS = New Recordset
      DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      DS.AddNew
      DS!CPF = TXT_CPF
      DS!AUXILIAR = "Adicionando"
      DS.UpdateBatch adAffectAll
      VarAcao = "Registro Novo"
      
      'btn_Limpar.Enabled = True
      cmdSalvar.Enabled = True
      cmdEditar.Enabled = False
      btn_Excluir.Enabled = False
      btn_Cartao.Enabled = False
      btn_Veiculo.Enabled = False
      
      Call LIMPAR
    End If
End If
End Sub
Private Sub TXT_E_MAIL1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_E_MAIL2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_LOGRADOURO_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_NOME_GUERRA_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_ORGAOEXPED_ID_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
