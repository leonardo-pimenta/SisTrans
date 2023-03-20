VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_CadFornecedor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SisTrans - Cadastro de Fornecedor"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7005
   ControlBox      =   0   'False
   Icon            =   "frm_CadFornecedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btn_Excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   2160
      Picture         =   "frm_CadFornecedor.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Excluir"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton btn_Salvar 
      Caption         =   "Sal&var"
      Enabled         =   0   'False
      Height          =   855
      Left            =   240
      Picture         =   "frm_CadFornecedor.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Salvar"
      Top             =   5640
      Width           =   855
   End
   Begin MSMask.MaskEdBox txt_CNPJ 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   12648447
      MaxLength       =   16
      Mask            =   "99999999/9999-99"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton btn_Cartao 
      Caption         =   "&Cartão"
      Enabled         =   0   'False
      Height          =   855
      Left            =   3120
      Picture         =   "frm_CadFornecedor.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Editar"
      Top             =   5640
      Width           =   855
   End
   Begin VB.Frame Frame5 
      Caption         =   "Lista de motoristas deste fornecedor:"
      Height          =   1575
      Left            =   240
      TabIndex        =   21
      Top             =   3840
      Width           =   6615
      Begin MSDataGridLib.DataGrid dbg_Listagem 
         Height          =   1215
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   4210688
         ForeColor       =   16777215
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
            MarqueeStyle    =   4
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton btn_Editar 
      Caption         =   "E&ditar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   1200
      Picture         =   "frm_CadFornecedor.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Editar"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton btn_Sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   6000
      Picture         =   "frm_CadFornecedor.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton btn_Veiculo 
      Caption         =   "Veí&culo"
      Enabled         =   0   'False
      Height          =   855
      Left            =   5040
      Picture         =   "frm_CadFornecedor.frx":1854
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Editar"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton btn_Motorista 
      Caption         =   "&Motorista"
      Enabled         =   0   'False
      Height          =   855
      Left            =   4080
      Picture         =   "frm_CadFornecedor.frx":1C96
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Editar"
      Top             =   5640
      Width           =   855
   End
   Begin VB.Frame fra_Cadastro 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5655
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   6735
      Begin VB.TextBox txt_NomeFornec 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   4695
      End
      Begin VB.Frame Frame4 
         Caption         =   "Complemento:"
         Height          =   1455
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   6615
         Begin VB.TextBox txt_Email 
            Height          =   285
            Left            =   3720
            TabIndex        =   12
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox txt_Responsavel 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   6135
         End
         Begin MSMask.MaskEdBox txt_Telefone 
            Height          =   300
            Left            =   120
            TabIndex        =   10
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   13
            Mask            =   "(99)9999-9999"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt_Fax 
            Height          =   300
            Left            =   1920
            TabIndex        =   11
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   13
            Mask            =   "(99)9999-9999"
            PromptChar      =   "_"
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Telefone:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Width           =   675
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
            Height          =   195
            Left            =   1920
            TabIndex        =   32
            Top             =   840
            Width           =   300
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "E-mail:"
            Height          =   195
            Left            =   3720
            TabIndex        =   31
            Top             =   840
            Width           =   465
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Responsável:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Height          =   135
         Left            =   120
         TabIndex        =   23
         Top             =   5400
         Width           =   6615
      End
      Begin VB.Frame Frame2 
         Caption         =   "Endereço:"
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   6615
         Begin VB.ComboBox txt_UF 
            Height          =   315
            ItemData        =   "frm_CadFornecedor.frx":20D8
            Left            =   3120
            List            =   "frm_CadFornecedor.frx":2130
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txt_Cidade 
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   1080
            Width           =   2895
         End
         Begin VB.TextBox txt_Logradouro 
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   3855
         End
         Begin VB.TextBox txt_Bairro 
            Height          =   285
            Left            =   4080
            TabIndex        =   4
            Top             =   480
            Width           =   2295
         End
         Begin MSMask.MaskEdBox txt_CEP 
            Height          =   300
            Left            =   4920
            TabIndex        =   7
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   9
            Mask            =   "99999-999"
            PromptChar      =   "_"
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro (rua, avenida, alameda, etc.) :"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   2970
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Left            =   3120
            TabIndex        =   28
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   840
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Left            =   4080
            TabIndex        =   26
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
            Height          =   195
            Left            =   4920
            TabIndex        =   25
            Top             =   840
            Width           =   360
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Fornecedor:"
         Height          =   195
         Left            =   1920
         TabIndex        =   24
         Top             =   120
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frm_CadFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_tabFornecedor As Recordset
Dim rs_Listagem As Recordset
Dim rs_Cartao As Recordset
Private Sub btn_Cartao_Click()

'Envia os dados do form atual para o de Cartao p/
'ser feitas as deidas operacoes

vgl_TipoResponsavel = "CNPJ"
frm_Cartao_Emitir.txt_Codigo = txt_CNPJ.Text
frm_Cartao_Emitir.txt_Descricao = txt_NomeFornec.Text

frm_Cartao_Emitir.Show
   
End Sub
Private Sub btn_Editar_Click()

fra_Cadastro.Enabled = True
txt_NomeFornec.SetFocus

btn_Salvar.Enabled = True
btn_Editar.Enabled = False
btn_Excluir.Enabled = False

btn_Cartao.Enabled = False
btn_Motorista.Enabled = False
btn_Veiculo.Enabled = False

End Sub
Private Sub btn_Motorista_Click()

'2 - abre em modo de adição
frm_CadMotoristaFornec.varModoAbrir = 2
frm_CadMotoristaFornec.varCNPJ = txt_CNPJ.Text
frm_CadMotoristaFornec.txt_NomeFornec = txt_NomeFornec.Text

frm_CadMotoristaFornec.Show

End Sub
Private Sub btn_Sair_Click()
Unload Me
End Sub

Private Sub btn_Salvar_Click()

If txt_CNPJ.Text <> "________/____-__" Then

'Abre a tabela verifica se a senha esta correta, caso sim edita-a
Set rs_tabFornecedor = New Recordset

With rs_tabFornecedor

'''''
    .Open "select * from tab_trans_fornecedor where cnpj ='" & txt_CNPJ.Text & "'", cnConexao, adOpenStatic, adLockOptimistic

    If .RecordCount = 0 Then
        .AddNew
        !cnpj = txt_CNPJ.Text
    End If
        
    !NOME = txt_NomeFornec.Text
    !responsavel = txt_Responsavel.Text
    !LOGRADOURO = TXT_LOGRADOURO.Text
    !BAIRRO = TXT_BAIRRO.Text
    !CIDADE = TXT_CIDADE.Text
    !uf = txt_UF.Text
    !CEP = TXT_CEP.Text
    !telefone = txt_Telefone.Text
    !fax = txt_Fax.Text
    !e_mail = txt_Email.Text
    
    .UpdateBatch adAffectAll
    MsgBox "Arquivo salvo.", vbOKOnly + vbInformation, "SisTrans"
    .Close

End With

End If

Call prc_LimparCampos
txt_CNPJ.Mask = "                "
txt_CNPJ.Mask = "99999999/9999-99"

End Sub

Private Sub btn_Veiculo_Click()

frm_CadVeiculo.vgl_Responsavel = 2

frm_CadVeiculo.opt_Fornecedor.Value = True
frm_CadVeiculo.txt_Responsavel = txt_CNPJ.Text
frm_CadVeiculo.txt_Descricao = txt_NomeFornec.Text

frm_CadVeiculo.Show

End Sub
Private Sub btn_Excluir_Click()

On Error GoTo Error

If txt_CNPJ.Text <> "___-____" Then
   If MsgBox("Deseja excluir o registro?", vbYesNo + vbQuestion, "SisTrans") = vbYes Then
      SQL = "delete * from tab_Trans_fornecedor where cnpj ='" & txt_CNPJ.Text & "'"
      Set DS = New Recordset
      DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      MsgBox "Registo Deletado.", vbOKOnly + vbInformation, "SisTrans"
   Else
      MsgBox "Operação cancelada.", vbOKOnly + vbInformation, "SisTrans"
      txt_CNPJ.SetFocus
      Exit Sub
   End If
   Call Form_Load
End If
Exit Sub
Error:
MsgBox "Não é possível excluir o registro, ele pode fazer parte de um ou mais relacionamentos.Ex.: multa x veículo, veículo x proprietário. Caso seja realmente necessária a exclusão, contate o administrador do Banco de Dados.", vbOKOnly + vbInformation, "SisTrans"
btn_Sair.SetFocus
    
End Sub
Private Sub dbg_Listagem_DblClick()

'1 - abre em modo normal
frm_CadMotoristaFornec.varModoAbrir = 1
frm_CadMotoristaFornec.varCodMotorista = rs_Listagem!cod_motorista
frm_CadMotoristaFornec.txt_NomeFornec = txt_NomeFornec.Text

frm_CadMotoristaFornec.Show

End Sub

Private Sub dbg_Listagem_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

    '1 - abre em modo normal
    frm_CadMotoristaFornec.varModoAbrir = 1
    frm_CadMotoristaFornec.varCodMotorista = rs_Listagem!cod_motorista
    
    frm_CadMotoristaFornec.Show
    
End If
End Sub
Private Sub Form_Activate()
If txt_CNPJ.Text <> "" Then
    Set rs_Listagem = New Recordset
    rs_Listagem.Open "select Cod_Motorista, cnpj_fornec, Nome_Motorista as Nome, Identidade FROM tab_Trans_Motorista_Fornecedor WHERE cnpj_fornec ='" & txt_CNPJ.Text & "' order by Nome_Motorista ", cnConexao, adOpenStatic, adLockOptimistic
    
    x = rs_Listagem.RecordCount
    
    Set dbg_Listagem.DataSource = rs_Listagem

    With dbg_Listagem
        .Columns(0).Visible = False
        .Columns(1).Visible = False
        .Columns(2).Width = 5000
        .Columns(3).Width = 2000
    End With
End If
End Sub
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Call prc_LimparCampos
txt_CNPJ.Mask = "                "
txt_CNPJ.Mask = "99999999/9999-99"

txt_UF.ListIndex = 0

    btn_Salvar.Enabled = False
btn_Editar.Enabled = False
btn_Excluir.Enabled = False
btn_Cartao.Enabled = False
btn_Motorista.Enabled = False
btn_Veiculo.Enabled = False

End Sub
Private Sub TXT_BAIRRO_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_CIDADE_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txt_CNPJ_GotFocus()

Call prc_LimparCampos

fra_Cadastro.Enabled = False

btn_Salvar.Enabled = False
btn_Editar.Enabled = False

btn_Cartao.Enabled = False
btn_Motorista.Enabled = False
btn_Veiculo.Enabled = False

End Sub
Private Sub txt_CNPJ_LostFocus()

If txt_CNPJ.Text = "________/____-__" Then Exit Sub

Set rs_tabFornecedor = New Recordset
With rs_tabFornecedor
    .Open "select * from tab_trans_fornecedor where cnpj ='" & txt_CNPJ.Text & "'", cnConexao, adOpenStatic, adLockOptimistic

    If .RecordCount = 1 Then
    
        txt_NomeFornec.Text = !NOME
        txt_Responsavel.Text = !responsavel
        TXT_LOGRADOURO.Text = !LOGRADOURO
        TXT_BAIRRO.Text = !BAIRRO
        TXT_CIDADE.Text = !CIDADE
        txt_UF.Text = !uf
        TXT_CEP.Text = !CEP
        txt_Telefone.Text = !telefone
        txt_Fax.Text = !fax
        txt_Email.Text = !e_mail
        
        Call Form_Activate
        
        btn_Editar.Enabled = True
        btn_Excluir.Enabled = True
        
        btn_Cartao.Enabled = True
        btn_Motorista.Enabled = True
        btn_Veiculo.Enabled = True
    Else
        fra_Cadastro.Enabled = True
        txt_NomeFornec.SetFocus
        
        btn_Salvar.Enabled = True
        btn_Editar.Enabled = False
        btn_Excluir.Enabled = False
        
        btn_Cartao.Enabled = False
        btn_Motorista.Enabled = False
        btn_Veiculo.Enabled = False
        Call prc_LimparCampos
    End If
End With
End Sub
Private Sub prc_LimparCampos()
txt_NomeFornec.Text = ""
txt_Responsavel.Text = ""
TXT_LOGRADOURO.Text = ""
TXT_BAIRRO.Text = ""
TXT_CIDADE.Text = ""
txt_UF.ListIndex = 0
TXT_CEP.Mask = "         "
TXT_CEP.Mask = "99999-999"
txt_Telefone.Mask = "             "
txt_Telefone.Mask = "(99)9999-9999"
txt_Fax.Mask = "             "
txt_Fax.Mask = "(99)9999-9999"
txt_Email.Text = ""
Set dbg_Listagem.DataSource = Nothing
End Sub
Private Sub TXT_LOGRADOURO_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txt_NomeFornec_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txt_Responsavel_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txt_UF_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
