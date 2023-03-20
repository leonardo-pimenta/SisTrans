VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.MDIForm frm_Principal 
   BackColor       =   &H80000009&
   Caption         =   "SisTrans - Sistema de Trânsito do Complexo do Com1ºDN"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11070
   Icon            =   "frm_Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frm_Principal.frx":0442
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport 
      Left            =   6480
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5280
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgMenu 
      Left            =   7320
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Principal.frx":2D3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Principal.frx":318E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Principal.frx":35E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Principal.frx":3A36
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Principal.frx":3E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Principal.frx":42DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Principal.frx":4732
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Principal.frx":4B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Principal.frx":4FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Principal.frx":542E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Principal.frx":5882
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Principal.frx":5CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Principal.frx":612A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Principal.frx":657E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Principal.frx":69D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Principal.frx":6E26
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMenuEquip 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   1005
      ButtonWidth     =   1614
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgMenu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Veículo"
            Object.ToolTipText     =   "Cadastro de veículos"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Multar"
            Object.ToolTipText     =   "Cadastra multas "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pessoa"
            Object.ToolTipText     =   "Cadsatro de pessoa"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fornecedor"
            Object.ToolTipText     =   "Cadstro de fornecedores"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "OM"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cartão"
            Object.ToolTipText     =   "Efetuar logoff de seu usuário"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Provisório"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cartao-OM"
            Object.ToolTipText     =   "Acesso apenas ao CGS"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "CGS"
            Object.ToolTipText     =   "Sair do sistema"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Logoff"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Object.ToolTipText     =   "Sair do sistema"
            ImageIndex      =   16
         EndProperty
      EndProperty
      BorderStyle     =   1
      MouseIcon       =   "frm_Principal.frx":727A
   End
   Begin VB.Menu mnu_Arquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnu_CGS 
         Caption         =   "C&GS"
         Begin VB.Menu Lista_Multa 
            Caption         =   "Listar multas a partir de uma data"
         End
         Begin VB.Menu Lista_requerimento 
            Caption         =   "Listar multas com requerimento"
         End
      End
      Begin VB.Menu tr11 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Logof 
         Caption         =   "Logof&f"
         Shortcut        =   ^G
      End
      Begin VB.Menu tr 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Sair 
         Caption         =   "Sai&r"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnu_Cartao 
      Caption         =   "&Cartão"
      Begin VB.Menu mnu_Veiculo 
         Caption         =   "&Veículo"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu_Proprietario 
         Caption         =   "Proprietário"
         Begin VB.Menu mnu_Pessoa 
            Caption         =   "&Pessoa"
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnu_OM 
            Caption         =   "&OM"
            Enabled         =   0   'False
            Shortcut        =   ^O
         End
         Begin VB.Menu mnu_ 
            Caption         =   "&Fornecedor"
            Shortcut        =   ^F
         End
      End
   End
   Begin VB.Menu mnu_Multa 
      Caption         =   "&Multa"
      Begin VB.Menu mnu_Multar 
         Caption         =   "&Multar"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnu_Requerimento 
         Caption         =   "&Requerer ao CGS"
         Shortcut        =   {F6}
      End
      Begin VB.Menu Detran 
         Caption         =   "&Multas ao DETRAN"
         Begin VB.Menu com72h 
            Caption         =   "&Com 72h"
         End
         Begin VB.Menu Todoss 
            Caption         =   "&Todos"
         End
      End
   End
   Begin VB.Menu mnu_Relatorios 
      Caption         =   "&Relatórios"
      Begin VB.Menu mnu_ReqCGS 
         Caption         =   "&Multas c/ Requerimento"
      End
      Begin VB.Menu List_SemReq 
         Caption         =   "M&ultas sem Requerimento"
      End
      Begin VB.Menu mnu_ListDeferidos 
         Caption         =   "&Listar Deferidos"
      End
      Begin VB.Menu List_indeferido 
         Caption         =   "L&istar Indeferido "
      End
      Begin VB.Menu nu_ListCartao 
         Caption         =   "&Listagem de Cartão"
         Begin VB.Menu proprietario 
            Caption         =   "&Proprietário"
         End
         Begin VB.Menu Extraviado 
            Caption         =   "&Extraviado"
         End
         Begin VB.Menu Ano 
            Caption         =   "P&or Ano"
         End
         Begin VB.Menu Veiculo 
            Caption         =   "&Número de cartão"
         End
         Begin VB.Menu Todos 
            Caption         =   "&Todos"
         End
      End
      Begin VB.Menu ListVeiculo 
         Caption         =   "&Listagem de Automóvel"
         Begin VB.Menu pessoaa 
            Caption         =   "&Pessoa"
         End
         Begin VB.Menu OMm 
            Caption         =   "&OM"
         End
         Begin VB.Menu Fornecedorr 
            Caption         =   "&Fornecedor"
         End
      End
      Begin VB.Menu moto 
         Caption         =   "Lis&tagem de Motocicleta"
         Begin VB.Menu pessoa 
            Caption         =   "&Pessoa"
         End
         Begin VB.Menu OM 
            Caption         =   "&OM"
         End
         Begin VB.Menu Fornecedor 
            Caption         =   "&Fornecedor"
         End
      End
      Begin VB.Menu mnu_MotorFornec 
         Caption         =   "Motorista/Fornec."
         Begin VB.Menu mnu_Todos_MotorFornec 
            Caption         =   "&Todos"
         End
         Begin VB.Menu mnu_Unit_MotorFornec 
            Caption         =   "&Unitário"
         End
      End
   End
   Begin VB.Menu mnu_Manutencao 
      Caption         =   "Ma&nutenção"
      Begin VB.Menu mnu_Usuário 
         Caption         =   "&Usuário"
         Begin VB.Menu mnu_InclirUsuario 
            Caption         =   "&Incluir novo usuário"
         End
         Begin VB.Menu mnu_AlterarSenha 
            Caption         =   "&Alterar Senha"
         End
      End
      Begin VB.Menu mnu_MarcaModelo 
         Caption         =   "&Marca/Modelo"
      End
      Begin VB.Menu mnu_Especialidade 
         Caption         =   "&Especialidade"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnu_Ajuda 
      Caption         =   "A&juda"
      Begin VB.Menu mnu_Conteudo 
         Caption         =   "&Conteúdo"
         Shortcut        =   {F1}
      End
      Begin VB.Menu tr7 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Sobre 
         Caption         =   "&Sobre o SisTrans"
      End
   End
End
Attribute VB_Name = "frm_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_Relatorio As Recordset
Dim var_path As String
Private Sub Consultar_Click()
frm_Cartao_Emitir.Show
End Sub
Private Sub Ano_Click()
'--------------------------
SQL = "Delete * from Tab_Trans_Aux_Cartao"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
'--------------------------
var_Resposta = InputBox("Digite o ano. Ex.:(2003)", "Lista de Multas")
If var_Resposta = "" Then Exit Sub
var_SQL = "select * FROM cons_cartao"
var_SQL = var_SQL + " where ano ='" & var_Resposta & "'"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
    DS.MoveFirst
    Do While Not DS.EOF
       SQL = "select * from Tab_Trans_Aux_Cartao"
       Set ds1 = New Recordset
       ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
       ds1.AddNew
       ds1!nr_cartao = DS(0)
       ds1!Ano = DS(1)
       ds1!TARJA = DS(2)
       ds1!Veiculo = DS(3)
       ds1!DT_EMISSAO = DS(4)
       ds1!COD_OM = DS(5)
       ds1!CPF_pessoa = DS(6)
       ds1!cnpj_fornecedor = DS(7)
       ds1!autorizacao = DS(8)
       ds1!data_validade = DS(9)
       ds1!Extraviado = DS(10)
       ds1!Motivo = DS(11)
       ds1!NOME = DS(12)
       ds1.UpdateBatch adAffectAll
       DS.MoveNext
    Loop
    CD.Flags = &H40
    CD.ShowPrinter
    CrystalReport.DataFiles(0) = (xpath)
    CrystalReport.ReportFileName = App.Path + "\cartao.rpt"
    CrystalReport.Action = 1
End If
End Sub
Private Sub Com72h_Click()
On Error GoTo Error

SQL = "delete * from Tab_Trans_Aux_Multa_Todos_Dados"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

Dim Var_Detran As String
Dim Var72h As Date
Dim var144h As Date
Var_Detran = "SIM"


var_SQL = "select * FROM cns_trans_multa_todos_dados"
var_SQL = var_SQL + " where dt_infr > " & Date - 4
var_SQL = var_SQL + " and detran = '" & Var_Detran & "'"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
    DS.MoveFirst
    Do While Not DS.EOF
       SQL = "select * from Tab_Trans_Aux_Multa_Todos_Dados"
       Set ds1 = New Recordset
       ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
       ds1.AddNew
       ds1!nr_talao_infr = DS(0)
       ds1!dt_infr = DS(1)
       ds1!hr_infr = DS(2)
       ds1!Local = DS(3)
       ds1!req_CGS = DS(4)
       ds1!deferido = DS(5)
       ds1!Detran = DS(6)
       ds1!Tab_Trans_Multa_placa = DS(7)
       ds1!nome_requerente = DS(8)
       ds1!rg_requerente = DS(9)
       ds1!orgao_exp = DS(10)
       ds1!Dt_Listagem_Detran = DS(11)
       ds1!Tab_Trans_veiculo_Placa = DS(12)
       ds1!Cor_Pred = DS(13)
       ds1!modelo = DS(14)
       ds1!MARCA = DS(15)
       ds1!desc_tipo_multa = DS(16)
       ds1.UpdateBatch adAffectAll
       DS.MoveNext
     Loop
     CD.Flags = &H40
     CD.ShowPrinter
     CrystalReport.DataFiles(0) = (xpath)
     CrystalReport.ReportFileName = App.Path + "\Detran.rpt"
     CrystalReport.Action = 1
End If
Error:
Exit Sub
End Sub
Private Sub Enviar_Detran_Click()
Dim var_deferido, var_Traco, var_Date As String
var_Traco = "-"
var_deferido = "DEFERIDO"
var_Date = Str(Date)
SQL = "select * FROM cns_Trans_multa_todos_dados "
SQL = SQL + " where deferido='" & var_Traco & "'"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set frm_Lista_multas.dbg_Listagem.DataSource = DS
With frm_Lista_multas.dbg_Listagem
     .Columns(0).Visible = False
     .Columns(4).Visible = False
     .Columns(6).Visible = False
     .Columns(9).Visible = False
     .Columns(10).Visible = False
     .Columns(11).Visible = False
     .Columns(12).Visible = False
     .Columns(13).Visible = False
     .Columns(14).Visible = False
     .Columns(15).Visible = False
     .Columns(1).Width = 1000
     .Columns(2).Width = 750
End With
frm_Lista_multas.Show
End Sub
Private Sub Extraviado_Click()
SQL = "Delete * from Tab_Trans_Aux_Cartao"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

var_Resposta = "Sim"
var_SQL = "select * FROM cons_cartao"
var_SQL = var_SQL + " where extraviado ='" & var_Resposta & "'"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
    DS.MoveFirst
    Do While Not DS.EOF
       SQL = "select * from Tab_Trans_Aux_Cartao"
       Set ds1 = New Recordset
       ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
       ds1.AddNew
       ds1!nr_cartao = DS(0)
       ds1!Ano = DS(1)
       ds1!TARJA = DS(2)
       ds1!Veiculo = DS(3)
       ds1!DT_EMISSAO = DS(4)
       ds1!COD_OM = DS(5)
       ds1!CPF_pessoa = DS(6)
       ds1!cnpj_fornecedor = DS(7)
       ds1!autorizacao = DS(8)
       ds1!data_validade = DS(9)
       ds1!Extraviado = DS(10)
       ds1!Motivo = DS(11)
       ds1!NOME = DS(12)
       ds1.UpdateBatch adAffectAll
       DS.MoveNext
     Loop
     CD.Flags = &H40
     CD.ShowPrinter
     CrystalReport.DataFiles(0) = (xpath)
     CrystalReport.ReportFileName = App.Path + "\Extraviado.rpt"
     CrystalReport.Action = 1
End If
End Sub
Private Sub Fornecedor_Click()
SQL = "delete * from Tab_Trans_Aux_veiculo_FORNECEDOR"
Set ds1 = New Recordset
ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

var_Resposta = InputBox("Entre com o CNPJ da empresa", "Lista de Multas")
If var_Resposta = "" Then Exit Sub
var_SQL = "select distinct placa,modelo,marca,tipo,nome  FROM cns_trans_veiculo_fornec"
var_SQL = var_SQL + " where cnpj ='" & var_Resposta & "'"
var_SQL = var_SQL + " and tipo = 'MOTOCICLETA'"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
   DS.MoveFirst
   Do While Not DS.EOF
      SQL = "select * from Tab_Trans_Aux_veiculo_FORNECEDOR"
      Set ds1 = New Recordset
      ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      ds1.AddNew
      ds1!placa = DS(0)
      ds1!modelo = DS(1)
      ds1!MARCA = DS(2)
      ds1!TIPO = DS(3)
      ds1!NOME = DS(4)
      ds1.UpdateBatch adAffectAll
      DS.MoveNext
   Loop
   CD.Flags = &H40
   CD.ShowPrinter
   CrystalReport.DataFiles(0) = (xpath)
   CrystalReport.ReportFileName = App.Path + "\Veiculo_Fornecedor.rpt"
   CrystalReport.Action = 1
End If
End Sub

Private Sub Fornecedorr_Click()
SQL = "delete * from Tab_Trans_Aux_veiculo_FORNECEDOR"
Set ds1 = New Recordset
ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

var_Resposta = InputBox("Entre com o CNPJ da empresa", "Lista de Multas")
If var_Resposta = "" Then Exit Sub
var_SQL = "select distinct placa,modelo,marca,tipo,nome  FROM cns_trans_veiculo_fornec"
var_SQL = var_SQL + " where cnpj ='" & var_Resposta & "'"
var_SQL = var_SQL + " and tipo = 'AUTOMÓVEL'"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
   DS.MoveFirst
   Do While Not DS.EOF
      SQL = "select * from Tab_Trans_Aux_veiculo_FORNECEDOR"
      Set ds1 = New Recordset
      ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      ds1.AddNew
      ds1!placa = DS(0)
      ds1!modelo = DS(1)
      ds1!MARCA = DS(2)
      ds1!TIPO = DS(3)
      ds1!NOME = DS(4)
      ds1.UpdateBatch adAffectAll
      DS.MoveNext
   Loop
   CD.Flags = &H40
   CD.ShowPrinter
   CrystalReport.DataFiles(0) = (xpath)
   CrystalReport.ReportFileName = App.Path + "\Veiculo_Fornecedor.rpt"
   CrystalReport.Action = 1
End If
End Sub

Private Sub List_indeferido_Click()
SQL = "delete * from Tab_Trans_Aux_Multa_Todos_Dados"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

Dim var_resposta_ini As Date
Dim var_resposta_fim As Date

x = InputBox("Digite a data inicial. Ex.:(25/12/2000)", "Lista de Multas")
If IsDate(x) Then
   var_resposta_ini = x
Else
   Resp = MsgBox("Data inválida", vbCritical)
   Exit Sub
End If

y = InputBox("Digite a data inicial. Ex.:(25/12/2000)", "Lista de Multas")
If IsDate(y) Then
   var_resposta_fim = y
Else
   Resp = MsgBox("Data inválida", vbCritical)
   Exit Sub
End If

var_deferido = "INDEFERIDO"

var_SQL = "SELECT *"
var_SQL = var_SQL + " From cns_trans_multa_todos_dados"
var_SQL = var_SQL + " WHERE (((cns_trans_multa_todos_dados.Dt_Infr) Between #" & var_resposta_ini & "# And #" & var_resposta_fim & "#))"
var_SQL = var_SQL + " and deferido = '" & var_deferido & "'"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic

If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
    SQL = "select * from Tab_Trans_Aux_Multa_Todos_Dados"
    Set ds1 = New Recordset
    ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
    DS.MoveFirst
    Do While Not DS.EOF
       ds1.AddNew
       ds1!nr_talao_infr = DS(0)
       ds1!dt_infr = DS(1)
       ds1!hr_infr = DS(2)
       ds1!Local = DS(3)
       ds1!req_CGS = DS(4)
       ds1!deferido = DS(5)
       ds1!Detran = DS(6)
       ds1!Tab_Trans_Multa_placa = DS(7)
       ds1!nome_requerente = DS(8)
       ds1!rg_requerente = DS(9)
       ds1!orgao_exp = DS(10)
       ds1!Dt_Listagem_Detran = DS(11)
       ds1!Tab_Trans_veiculo_Placa = DS(12)
       ds1!Cor_Pred = DS(13)
       ds1!modelo = DS(14)
       ds1!MARCA = DS(15)
       ds1!desc_tipo_multa = DS(16)
       ds1.UpdateBatch adAffectAll
       DS.MoveNext
     Loop
     CD.Flags = &H40
     CD.ShowPrinter
     CrystalReport.DataFiles(0) = (xpath)
     CrystalReport.ReportFileName = App.Path + "\deferimento.rpt"
     CrystalReport.Action = 1
End If
End Sub

Private Sub List_SemReq_Click()
SQL = "delete * from Tab_Trans_Aux_Multa_Todos_Dados"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

Dim var_resposta_ini As Date
Dim var_resposta_fim As Date

x = InputBox("Digite a data inicial. Ex.:(25/12/2000)", "Lista de Multas")
If IsDate(x) Then
   var_resposta_ini = x
Else
   Resp = MsgBox("Data inválida", vbCritical)
   Exit Sub
End If

y = InputBox("Digite a data inicial. Ex.:(25/12/2000)", "Lista de Multas")
If IsDate(y) Then
   var_resposta_fim = y
Else
   Resp = MsgBox("Data inválida", vbCritical)
   Exit Sub
End If


var_SQL = "SELECT * FROM CNS_DATA"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
    DS.MoveFirst
    Do While Not DS.EOF
       SQL = "select * from Tab_Trans_Aux_Multa_Todos_Dados"
       Set ds1 = New Recordset
       ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
       ds1.AddNew
       ds1!nr_talao_infr = DS(0)
       ds1!dt_infr = DS(1)
       ds1!hr_infr = DS(2)
       ds1!Local = DS(3)
       ds1!req_CGS = DS(4)
       ds1!deferido = DS(5)
       ds1!Detran = DS(6)
       ds1!Tab_Trans_Multa_placa = DS(7)
       ds1!nome_requerente = DS(8)
       ds1!rg_requerente = DS(9)
       ds1!orgao_exp = DS(10)
       ds1!Dt_Listagem_Detran = DS(11)
       ds1!Tab_Trans_veiculo_Placa = DS(12)
       ds1!Cor_Pred = DS(13)
       ds1!modelo = DS(14)
       ds1!MARCA = DS(15)
       ds1!desc_tipo_multa = DS(16)
       ds1.UpdateBatch adAffectAll
       DS.MoveNext
     Loop
     CD.Flags = &H40
     CD.ShowPrinter
     CrystalReport.DataFiles(0) = (xpath)
     CrystalReport.ReportFileName = App.Path + "\Requerimento.rpt"
     CrystalReport.Action = 1
End If
End Sub
Private Sub Lista_Multa_Click()

Dim var_resposta_ini As Date
Dim var_resposta_fim As Date

x = InputBox("Digite a data inicial. Ex.:(25/12/2000)", "Lista de Multas")
If IsDate(x) Then
   var_resposta_ini = x
Else
   Resp = MsgBox("Data inválida", vbCritical)
   Exit Sub
End If

y = InputBox("Digite a data inicial. Ex.:(25/12/2000)", "Lista de Multas")
If IsDate(y) Then
   var_resposta_fim = y
Else
   Resp = MsgBox("Data inválida", vbCritical)
   Exit Sub
End If

var_SQL = "SELECT *"
var_SQL = var_SQL + " From cns_trans_multa_todos_dados"
var_SQL = var_SQL + " WHERE (((cns_trans_multa_todos_dados.Dt_Infr) Between #" & var_resposta_ini & "# And #" & var_resposta_fim & "#))"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
   Set frm_ListaMultaCGS.dbg_Listagem.DataSource = DS
   frm_ListaMultaCGS.Show
   
End If
End Sub
Public Sub Lista_requerimento_Click()

var_deferido = "-"
SQL = "select * FROM cns_Trans_multa_todos_dados "
SQL = SQL + " WHERE (req_cgs <> '" & "'"
SQL = SQL + " AND deferido = '" & var_deferido & "')"
SQL = SQL + " AND dt_infr > " & Date - 4
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set frm_ListaMultaCGS.dbg_Listagem.DataSource = DS
frm_ListaMultaCGS.Show
If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
End If
End Sub
Private Sub MDIForm_Load()
frm_Principal.Caption = "SisTrans - Sistema de Trânsito do Complexo do Com1ºDN - Versão " & App.Major & "." & App.Minor & "." & App.Revision & " - Usuário: " & vgl_Responsavel
Call pcd_Usuario_Habilitar(vgl_Nivel)
vgl_X = Int(frm_Principal.Width) / 2
vgl_Y = Int(frm_Principal.Height) / 2
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
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

Private Sub mnu__Click()
frm_CadFornecedor.Show
End Sub
Private Sub mnu_AlterarSenha_Click()
frm_Senha_Alterar.Show
End Sub
Private Sub mnu_ConfEnvDetran_Click()
var_MSG = "Entre com a DATA da listagem de que se deseja confirmar o "
var_MSG = var_MSG + "envio ao DETRAN. Ex.: (25/12/2000) "
var_MSG = var_MSG + "NOTA: A data da listegem localiza-se no rodapé da impressão."

var_Resposta = InputBox(var_MSG, "Confirmar envio ao DETRAN")
If var_Resposta = "" Then Exit Sub

SQL = "UPDATE tab_trans_multa SET tab_trans_multa.Detran = 'SIM'"
SQL = SQL + " WHERE tab_trans_multa.Dt_Listagem_DETRAN='" & var_Resposta & "'"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

MsgBox "Confirmado o envio das multas listadas no dia " & var_Resposta, vbInformation, "SISTRANS"

End Sub

Private Sub mnu_ConfirEnvioDetran_Click()

var_Resposta = InputBox("Entre com a Data em que foi emitida a Listagem. Ex.:(25/12/2000)", "Lista de Multas")
If var_Resposta = "" Then Exit Sub

MsgBox "Será impresso a lista das multas que não foi  confirma p/ o envio ao DETRAN.", vbInformation + vbOKOnly, "SisTrans"

var_SQL = "select * FROM cns_Trans_multa_todos_dados "
var_SQL = var_SQL + "WHERE Dt_Listagem_Detran = '" & var_Resposta & "' AND Detran = 'NÃO'"

Set rs_Relatorio = New Recordset
rs_Relatorio.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic

Set rpt_MultaDetran.DataSource = rs_Relatorio
rpt_MultaDetran.Show

End Sub
Private Sub mnu_Conteudo_Click()
frm_Help.Show
End Sub
Private Sub mnu_DataDetran_Click()

var_Resposta = InputBox("Entre com a Data da Infração. Ex.:(25/12/2000)", "Lista de Multas")
If var_Resposta = "" Then Exit Sub

MsgBox "Será impresso a lista das multas que tiveram o requerimento de anulação indeferido pelo Sr CGS ou que não fizeram requerimento em até 48hs após a data de infração", vbInformation + vbOKOnly, "SisTrans"

var_SQL = "select * FROM cns_Trans_multa_todos_dados "
var_SQL = var_SQL + "WHERE dt_infr = '" & var_Resposta & "' AND Detran = 'SIM'"

Set rs_Relatorio = New Recordset
rs_Relatorio.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic

Set rpt_MultaDetran.DataSource = rs_Relatorio
rpt_MultaDetran.Show

End Sub
Private Sub mnu_Especialidade_Click()
frm_Especialidade_Subespecialidade.Show
End Sub

Private Sub mnu_InclirUsuario_Click()
frm_Senha_Incluir.Show
End Sub

Private Sub mnu_Logof_Click()
If MsgBox("Deseja efetuar logoff de " & vgl_Responsavel & "?", vbQuestion + vbYesNo, "Confirmação Saída") = vbYes Then
    Unload frm_Principal
    frm_Login.Show
End If
End Sub

Private Sub mnu_ListaDetran_Click()

Dim var_Dt_ListaDetran As String

'retorna a data de 48hs atras, tirando fim de semana e rotinas de domingo
var_Data = fun_DiasUteis(Date, 3)

var_Dt_ListaDetran = ""
var_path = "NAOFIG"

var_MSG = "Será impresso uma lista das multas CADASTRADAS NO DIA " & var_Data
var_MSG = var_MSG + " que tiveram o requerimento "
var_MSG = var_MSG + "de anulação indeferido pelo Sr CGS ou que não fizeram "
var_MSG = var_MSG + "requerimento em até 48hs após a data de infração. "
var_MSG = var_MSG + "Essa operação só poderá ser realizada UMA VEZ AO DIA."

MsgBox var_MSG, vbInformation + vbOKOnly, "SisTrans"

Set rs_Relatorio = New Recordset

'Consulta com as condicoes indicadas
var_deferido = "INDEFERIDO"

var_SQL = "select * FROM cns_Trans_multa_todos_dados "
var_SQL = var_SQL + " WHERE (deferido =  '" & var_deferido & "'"
var_SQL = var_SQL + " OR (req_cgs = '" & var_path & "'"
var_SQL = var_SQL + " And dt_infr = '" & var_Data & "')) "
var_SQL = var_SQL + " And Dt_Listagem_Detran = '" & var_Dt_ListaDetran & "'"
rs_Relatorio.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic

'verifica se há ou não multas p/ listar
If rs_Relatorio.RecordCount = 0 Then
    var_MSG = "Essa listagem já foi impressa "
    var_MSG = var_MSG + "ou não há multas cadastradas no dia " & var_Data
    MsgBox var_MSG, vbOKOnly + vbInformation, "SisTrans"
Else
    Set rpt_MultaDetran.DataSource = rs_Relatorio
    rpt_MultaDetran.Show
    
    'Atualiza a data de listagem para controle
    rs_Relatorio.MoveFirst
    For x = 1 To rs_Relatorio.RecordCount
        rs_Relatorio!Dt_Listagem_Detran = Str(Date)
        rs_Relatorio.UpdateBatch adAffectAll
        rs_Relatorio.MoveNext
    Next
End If
End Sub
Private Sub mnu_ListDeferidos_Click()
SQL = "delete * from Tab_Trans_Aux_Multa_Todos_Dados"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

Dim var_resposta_ini As Date
Dim var_resposta_fim As Date

x = InputBox("Digite a data inicial. Ex.:(25/12/2000)", "Lista de Multas")
If IsDate(x) Then
   var_resposta_ini = x
Else
   Resp = MsgBox("Data inválida", vbCritical)
   Exit Sub
End If

y = InputBox("Digite a data inicial. Ex.:(25/12/2000)", "Lista de Multas")
If IsDate(y) Then
   var_resposta_fim = y
Else
   Resp = MsgBox("Data inválida", vbCritical)
   Exit Sub
End If


var_deferido = "DEFERIDO"

var_SQL = "SELECT *"
var_SQL = var_SQL + " From cns_trans_multa_todos_dados"
var_SQL = var_SQL + " WHERE (((cns_trans_multa_todos_dados.Dt_Infr) Between #" & var_resposta_ini & "# And #" & var_resposta_fim & "#))"
var_SQL = var_SQL + " and deferido = '" & var_deferido & "'"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic

If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
    DS.MoveFirst
    Do While Not DS.EOF
       SQL = "select * from Tab_Trans_Aux_Multa_Todos_Dados"
       Set ds1 = New Recordset
       ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
       ds1.AddNew
       ds1!nr_talao_infr = DS(0)
       ds1!dt_infr = DS(1)
       ds1!hr_infr = DS(2)
       ds1!Local = DS(3)
       ds1!req_CGS = DS(4)
       ds1!deferido = DS(5)
       ds1!Detran = DS(6)
       ds1!Tab_Trans_Multa_placa = DS(7)
       ds1!nome_requerente = DS(8)
       ds1!rg_requerente = DS(9)
       ds1!orgao_exp = DS(10)
       ds1!Dt_Listagem_Detran = DS(11)
       ds1!Tab_Trans_veiculo_Placa = DS(12)
       ds1!Cor_Pred = DS(13)
       ds1!modelo = DS(14)
       ds1!MARCA = DS(15)
       ds1!desc_tipo_multa = DS(16)
       ds1.UpdateBatch adAffectAll
       DS.MoveNext
     Loop
     CD.Flags = &H40
     CD.ShowPrinter
     CrystalReport.DataFiles(0) = (xpath)
     CrystalReport.ReportFileName = App.Path + "\Deferimento.rpt"
     CrystalReport.Action = 1
End If
End Sub
Private Sub mnu_MarcaModelo_Click()
frm_Marca_Modelo.Show
End Sub
Private Sub mnu_Multar_Click()
frm_CadMulta.Show
End Sub
Private Sub mnu_OM_Click()
Frm_OM.Show
End Sub
Private Sub mnu_Pessoa_Click()
frm_CadPessoa.Show
End Sub
Private Sub mnu_ReqCGS_Click()
SQL = "delete * from Tab_Trans_Aux_Multa_Todos_Dados"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   
Dim var_resposta_ini As Date
Dim var_resposta_fim As Date

x = InputBox("Digite a data inicial. Ex.:(25/12/2000)", "Lista de Multas")
If IsDate(x) Then
   var_resposta_ini = x
Else
   Resp = MsgBox("Data inválida", vbCritical)
   Exit Sub
End If

y = InputBox("Digite a data inicial. Ex.:(25/12/2000)", "Lista de Multas")
If IsDate(y) Then
   var_resposta_fim = y
Else
   Resp = MsgBox("Data inválida", vbCritical)
   Exit Sub
End If


var_SQL = "SELECT *"
var_SQL = var_SQL + " From cns_trans_multa_todos_dados"
var_SQL = var_SQL + " WHERE (((cns_trans_multa_todos_dados.Dt_Infr) Between #" & var_resposta_ini & "# And #" & var_resposta_fim & "#))"
var_SQL = var_SQL + " and req_cgs <> '" & "'"

Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
    SQL = "select * from Tab_Trans_Aux_Multa_Todos_Dados"
    Set ds1 = New Recordset
    ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
    DS.MoveFirst
    Do While Not DS.EOF
       ds1.AddNew
       ds1!nr_talao_infr = DS(0)
       ds1!dt_infr = DS(1)
       ds1!hr_infr = DS(2)
       ds1!Local = DS(3)
       ds1!req_CGS = DS(4)
       ds1!deferido = DS(5)
       ds1!Detran = DS(6)
       ds1!Tab_Trans_Multa_placa = DS(7)
       ds1!nome_requerente = DS(8)
       ds1!rg_requerente = DS(9)
       ds1!orgao_exp = DS(10)
       ds1!Dt_Listagem_Detran = DS(11)
       ds1!Tab_Trans_veiculo_Placa = DS(12)
       ds1!Cor_Pred = DS(13)
       ds1!modelo = DS(14)
       ds1!MARCA = DS(15)
       ds1!desc_tipo_multa = DS(16)
       ds1.UpdateBatch adAffectAll
       DS.MoveNext
     Loop
     CD.Flags = &H40
     CD.ShowPrinter
     CrystalReport.DataFiles(0) = (xpath)
     CrystalReport.ReportFileName = App.Path + "\Requerimento.rpt"
     CrystalReport.Action = 1
End If
End Sub
Private Sub mnu_Requerimento_Click()
frm_CadMultaReq.Show
End Sub
Private Sub mnu_RotinaDomingo_Click()
frm_Rotina_Domingo.Show
End Sub
Private Sub mnu_Sair_Click()
If MsgBox("Deseja fechar o SisTrans?", vbQuestion + vbYesNo, "Com1DN") = vbYes Then
End
End If
End Sub
Private Sub mnu_Sobre_Click()
frm_Sobre.Show
End Sub
Private Sub mnu_Todos_MotorFornec_Click()
SQL = "delete * from Tab_Trans_Aux_Fornecedor"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set DS = New Recordset
With DS
    .Open "select * FROM cns_Trans_fornec_motorista", cnConexao, adOpenStatic, adLockOptimistic
    If .RecordCount = 0 Then
        MsgBox "Não há motoristas a listar.", vbInformation + vbOKOnly, "SisTrans"
        .Close
      Else
      DS.MoveFirst
      Do While Not DS.EOF
         SQL = "select * from Tab_Trans_Aux_Fornecedor"
         Set ds1 = New Recordset
         ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
         ds1.AddNew
         ds1!cnpj = DS(0)
         ds1!NOME = DS(1)
         ds1!telefone = DS(2)
         ds1!nome_motorista = DS(3)
         ds1!identidade = DS(4)
         ds1!OrgaoExped_Id = DS(5)
         ds1.UpdateBatch adAffectAll
         DS.MoveNext
       Loop
       CD.Flags = &H40
       CD.ShowPrinter
       CrystalReport.DataFiles(0) = (xpath)
       CrystalReport.ReportFileName = App.Path + "\Motorista.rpt"
       CrystalReport.Action = 1
    End If
End With
End Sub
Private Sub mnu_Unit_MotorFornec_Click()
SQL = "delete * from Tab_Trans_Aux_Fornecedor"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

var_Resposta = InputBox("Entre com o CNPJ do fornecedor.", "Lista de motoristas")
If var_Resposta = "" Then Exit Sub

Set DS = New Recordset
With DS
    .Open "select * FROM cns_Trans_fornec_motorista WHERE cnpj = '" & var_Resposta & "'", cnConexao, adOpenStatic, adLockOptimistic
    If .RecordCount = 0 Then
        MsgBox "Não há motoristas a listar.", vbInformation + vbOKOnly, "SisTrans"
        .Close
    Else
       DS.MoveFirst
       Do While Not DS.EOF
          SQL = "select * from Tab_Trans_Aux_Fornecedor"
          Set ds1 = New Recordset
          ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
          ds1.AddNew
          ds1!cnpj = DS(0)
          ds1!NOME = DS(1)
          ds1!telefone = DS(2)
          ds1!nome_motorista = DS(3)
          ds1!identidade = DS(4)
          ds1!OrgaoExped_Id = DS(5)
          ds1.UpdateBatch adAffectAll
          DS.MoveNext
        Loop
        CD.Flags = &H40
        CD.ShowPrinter
        CrystalReport.DataFiles(0) = (xpath)
        CrystalReport.ReportFileName = App.Path + "\Motorista.rpt"
        CrystalReport.Action = 1
    End If
End With
End Sub
Private Sub mnu_Veiculo_Click()
frm_CadVeiculo.vgl_Responsavel = 4
frm_CadVeiculo.Show
End Sub
Private Sub Provisorio_Click()
Call frm_Provisorio.Show
End Sub
Private Sub OM_Click()
SQL = "delete * from Tab_Trans_Aux_veiculo_OM"
Set ds1 = New Recordset
ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

var_Resposta = InputBox("Entre o código da OM", "Lista de Multas")
If var_Resposta = "" Then Exit Sub
var_SQL = "select * FROM cns_trans_veiculo_OM"
var_SQL = var_SQL + " where cod ='" & var_Resposta & "'"
var_SQL = var_SQL + " and tipo = 'MOTOCICLETA'"

Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
   DS.MoveFirst
   Do While Not DS.EOF
      SQL = "select * from Tab_Trans_Aux_veiculo_OM"
      Set ds1 = New Recordset
      ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      ds1.AddNew
      ds1!placa = DS(0)
      ds1!TIPO = DS(1)
      ds1!modelo = DS(2)
      ds1!MARCA = DS(3)
      ds1!cod = DS(4)
      ds1!SIGLA = DS(5)
      ds1!NOME = DS(6)
      ds1.UpdateBatch adAffectAll
      DS.MoveNext
   Loop
   CD.Flags = &H40
   CD.ShowPrinter
   CrystalReport.DataFiles(0) = (xpath)
   CrystalReport.ReportFileName = App.Path + "\Veiculo_OM.rpt"
   CrystalReport.Action = 1
End If

End Sub

Private Sub OMm_Click()
SQL = "delete * from Tab_Trans_Aux_veiculo_OM"
Set ds1 = New Recordset
ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

var_Resposta = InputBox("Entre o código da OM", "Lista de Multas")
If var_Resposta = "" Then Exit Sub
var_SQL = "select * FROM cns_trans_veiculo_OM"
var_SQL = var_SQL + " where cod ='" & var_Resposta & "'"
var_SQL = var_SQL + " and tipo = 'AUTOMÓVEL'"

Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
   DS.MoveFirst
   Do While Not DS.EOF
      SQL = "select * from Tab_Trans_Aux_veiculo_OM"
      Set ds1 = New Recordset
      ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      ds1.AddNew
      ds1!placa = DS(0)
      ds1!TIPO = DS(1)
      ds1!modelo = DS(2)
      ds1!MARCA = DS(3)
      ds1!cod = DS(4)
      ds1!SIGLA = DS(5)
      ds1!NOME = DS(6)
      ds1.UpdateBatch adAffectAll
      DS.MoveNext
   Loop
   CD.Flags = &H40
   CD.ShowPrinter
   CrystalReport.DataFiles(0) = (xpath)
   CrystalReport.ReportFileName = App.Path + "\Veiculo_OM.rpt"
   CrystalReport.Action = 1
End If
End Sub

Private Sub pessoa_Click()
SQL = "Delete * from Tab_Trans_Aux_Veiculo_pessoa"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

var_Resposta = InputBox("Entre com o CPF da pessoa", "Lista de Multas")
If var_Resposta = "" Then Exit Sub
var_SQL = "select * FROM cns_trans_veiculo_pessoa"
var_SQL = var_SQL + " where cpf ='" & var_Resposta & "'"
var_SQL = var_SQL + " and tipo = 'MOTOCICLETA'"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
   DS.MoveFirst
   Do While Not DS.EOF
      SQL = "select * from Tab_Trans_Aux_veiculo_pessoa"
      Set ds1 = New Recordset
      ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      ds1.AddNew
      ds1!CPF = DS(0)
      ds1!NOME = DS(1)
      ds1!POST_GRAD_CAT_FUNC = DS(2)
      ds1!SIGLA = DS(3)
      ds1!placa = DS(4)
      ds1!modelo = DS(5)
      ds1!MARCA = DS(6)
      ds1!TIPO = DS(7)
      ds1.UpdateBatch adAffectAll
      DS.MoveNext
   Loop
   CD.Flags = &H40
   CD.ShowPrinter
   CrystalReport.DataFiles(0) = (xpath)
   CrystalReport.ReportFileName = App.Path + "\Veiculo_Pessoa.rpt"
   CrystalReport.Action = 1
End If
End Sub

Private Sub pessoaa_Click()
SQL = "Delete * from Tab_Trans_Aux_Veiculo_pessoa"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

var_Resposta = InputBox("Entre com o CPF da pessoa", "Lista de Multas")
If var_Resposta = "" Then Exit Sub
var_SQL = "select * FROM cns_trans_veiculo_pessoa"
var_SQL = var_SQL + " where cpf ='" & var_Resposta & "'"
var_SQL = var_SQL + " and tipo = 'AUTOMÓVEL'"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
   DS.MoveFirst
   Do While Not DS.EOF
      SQL = "select * from Tab_Trans_Aux_veiculo_pessoa"
      Set ds1 = New Recordset
      ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      ds1.AddNew
      ds1!CPF = DS(0)
      ds1!NOME = DS(1)
      ds1!POST_GRAD_CAT_FUNC = DS(2)
      ds1!SIGLA = DS(3)
      ds1!placa = DS(4)
      ds1!modelo = DS(5)
      ds1!MARCA = DS(6)
      ds1!TIPO = DS(7)
      ds1.UpdateBatch adAffectAll
      DS.MoveNext
   Loop
   CD.Flags = &H40
   CD.ShowPrinter
   CrystalReport.DataFiles(0) = (xpath)
   CrystalReport.ReportFileName = App.Path + "\Veiculo_Pessoa.rpt"
   CrystalReport.Action = 1
End If
End Sub
Private Sub proprietario_Click()
'-------------------------------------------------------
SQL = "Delete * from Tab_Trans_Aux_Cartao"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
'-------------------------------------------------------

var_Resposta = UCase(InputBox("Digite o nome da pessoa", "SisTrans"))
SQL = "select * FROM cons_cartao"
SQL = SQL + " where nome like '%" & var_Resposta & "%'"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
    DS.MoveFirst
    Do While Not DS.EOF
       SQL = "select * from Tab_Trans_Aux_Cartao"
       Set ds1 = New Recordset
       ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
       ds1.AddNew
       ds1!nr_cartao = DS(0)
       ds1!Ano = DS(1)
       ds1!TARJA = DS(2)
       ds1!Veiculo = DS(3)
       ds1!DT_EMISSAO = DS(4)
       ds1!COD_OM = DS(5)
       ds1!CPF_pessoa = DS(6)
       ds1!cnpj_fornecedor = DS(7)
       ds1!autorizacao = DS(8)
       ds1!data_validade = DS(9)
       ds1!Extraviado = DS(10)
       ds1!Motivo = DS(11)
       ds1!NOME = DS(12)
       ds1.UpdateBatch adAffectAll
       DS.MoveNext
    Loop
    CD.Flags = &H40
    CD.ShowPrinter
    CrystalReport.DataFiles(0) = (xpath)
    CrystalReport.ReportFileName = App.Path + "\Cartao.rpt"
    CrystalReport.Action = 1
End If

End Sub
Private Sub tlbMenuEquip_ButtonClick(ByVal Button As MSComctlLib.Button)


Select Case Button.Index
    Case 1
        Call mnu_Veiculo_Click
    Case 2
        frm_CadMulta.Show
    Case 4
        frm_CadPessoa.Show
    Case 5
        frm_CadFornecedor.Show
    Case 6
        Frm_OM.Show
    Case 8
        frm_Cartao_Emitir.Show
    Case 9
        var_respostaa = UCase(InputBox("Entre com o validade do cartão", "SisTrans"))
        If IsDate(var_respostaa) Then
           If var_respostaa = "" Then
              Exit Sub
           Else
              var_Resposta = UCase(InputBox("Entre com o CPF Ex.: 999.999.999-99", "SisTrans"))
              If var_Resposta = "" Then
                 Exit Sub
              End If
           End If
           SQL = "select * FROM TAB_GER_PESSOA,tab_trans_aux_posto_tarja"
           SQL = SQL + " where CPF='" & var_Resposta & "'"
           SQL = SQL + " and cod_postGrad_PostGradCatFunc=postoGrad"
           Set DS = New Recordset
           DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
           If DS.RecordCount = 0 Then
              MsgBox "CPF inválido!", vbOKOnly + vbInformation, "SisTrans"
              Exit Sub
           End If
           If DS!tarjacartao = "" Then
              var_MSG = "A graduação a que pertence o militar não possui permissão de se "
              var_MSG = var_MSG + "emitir um cartão de estacionamento ou a graduação a que "
              var_MSG = var_MSG + "pertence não foi cadastrada junto a uma tarja."
              MsgBox var_MSG, vbOKOnly + vbInformation, "SisTrans"
              Exit Sub
           ElseIf DS.RecordCount <> 0 Then
              Var_CPF = DS!CPF
              Var_nome = DS!NOME
              vgl_PostoGrad = DS!cod_postGrad_PostGradCatFunc
              vgl_TarjaCartao = DS!tarjacartao
              vgl_TipoResponsavel = "CPF"
              frm_Cartao_Emitir.Txt_Validade = var_respostaa
              frm_Cartao_Emitir.txt_Codigo = Var_CPF
              frm_Cartao_Emitir.txt_Descricao = Var_nome
              frm_Cartao_Emitir.Txt_Tarja = vgl_TarjaCartao
           Else
              MsgBox "CPF inválido ou não cadastrado!", vbInformation + vbOKOnly, "SisTrans"
           End If
        Else
           MsgBox "Data inválida", vbInformation + vbOKOnly, "SisTrans"
        End If
    Case 10
        vgl_placa = ""
        var_Resposta = UCase(InputBox("Entre com o código da OM", "SisTrans"))
        If var_Resposta = "" Then
           Exit Sub
        Else
           SQL = "select * FROM TAB_GER_OM"
           SQL = SQL + " where COD='" & var_Resposta & "'"
           Set DS = New Recordset
           DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
           If DS.RecordCount <> 0 Then
              Var_COD = DS!cod
              Var_nome = DS!NOME
              vgl_TipoResponsavel = "OM"
              frm_Cartao_Emitir.txt_Codigo = Var_COD
              frm_Cartao_Emitir.txt_Descricao = Var_nome
              frm_Cartao_Emitir.Txt_Tarja = "AZUL"
              vgl_Cor = "Azul"
           Else
              MsgBox "Código de OM inválido ou não cadastrado!", vbInformation + vbOKOnly, "SisTrans"
           End If
        End If
    Case 12
        Call Lista_requerimento_Click
    Case 13
        Call mnu_Logof_Click
    Case 15
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
        End
End Select
End Sub
Private Sub Todos_Click()
SQL = "Delete * from Tab_Trans_Aux_Cartao"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

var_SQL = "select * FROM cons_cartao"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
    DS.MoveFirst
    Do While Not DS.EOF
       SQL = "select * from Tab_Trans_Aux_Cartao"
       Set ds1 = New Recordset
       ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
       ds1.AddNew
       ds1!nr_cartao = DS(0)
       ds1!Ano = DS(1)
       ds1!TARJA = DS(2)
       ds1!Veiculo = DS(3)
       ds1!DT_EMISSAO = DS(4)
       ds1!COD_OM = DS(5)
       ds1!CPF_pessoa = DS(6)
       ds1!cnpj_fornecedor = DS(7)
       ds1!autorizacao = DS(8)
       ds1!data_validade = DS(9)
       ds1!Extraviado = DS(10)
       ds1!Motivo = DS(11)
       ds1!NOME = DS(12)
       ds1.UpdateBatch adAffectAll
       DS.MoveNext
     Loop
     CD.Flags = &H40
     CD.ShowPrinter
     CrystalReport.DataFiles(0) = (xpath)
     CrystalReport.ReportFileName = App.Path + "\cartao.rpt"
     CrystalReport.Action = 1
End If
End Sub

Private Sub Todoss_Click()
On Error GoTo Error

SQL = "delete * from Tab_Trans_Aux_Multa_Todos_Dados"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

Var_Detran = "SIM"

Dim var_resposta_ini As Date
Dim var_resposta_fim As Date

x = InputBox("Digite a data inicial. Ex.:(25/12/2000)", "Lista de Multas")
If IsDate(x) Then
   var_resposta_ini = x
Else
   Resp = MsgBox("Data inválida", vbCritical)
   Exit Sub
End If

y = InputBox("Digite a data inicial. Ex.:(25/12/2000)", "Lista de Multas")
If IsDate(y) Then
   var_resposta_fim = y
Else
   Resp = MsgBox("Data inválida", vbCritical)
   Exit Sub
End If

var_SQL = "SELECT *"
var_SQL = var_SQL + " From cns_trans_multa_todos_dados"
var_SQL = var_SQL + " WHERE (((cns_trans_multa_todos_dados.Dt_Infr) Between #" & var_resposta_ini & "# And #" & var_resposta_fim & "#))"
var_SQL = var_SQL + " and detran = '" & Var_Detran & "'"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
    MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
Else
DS.MoveFirst
    Do While Not DS.EOF
       SQL = "select * from Tab_Trans_Aux_Multa_Todos_Dados"
       Set ds1 = New Recordset
       ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
       ds1.AddNew
       ds1!nr_talao_infr = DS(0)
       ds1!dt_infr = DS(1)
       ds1!hr_infr = DS(2)
       ds1!Local = DS(3)
       ds1!req_CGS = DS(4)
       ds1!deferido = DS(5)
       ds1!Detran = DS(6)
       ds1!Tab_Trans_Multa_placa = DS(7)
       ds1!nome_requerente = DS(8)
       ds1!rg_requerente = DS(9)
       ds1!orgao_exp = DS(10)
       ds1!Dt_Listagem_Detran = DS(11)
       ds1!Tab_Trans_veiculo_Placa = DS(12)
       ds1!Cor_Pred = DS(13)
       ds1!modelo = DS(14)
       ds1!MARCA = DS(15)
       ds1!desc_tipo_multa = DS(16)
       ds1.UpdateBatch adAffectAll
       DS.MoveNext
     Loop
     CD.Flags = &H40
     CD.ShowPrinter
     CrystalReport.DataFiles(0) = (xpath)
     CrystalReport.ReportFileName = App.Path + "\Detran.rpt"
     CrystalReport.Action = 1
End If
Error:
 Exit Sub
End Sub

Private Sub veiculo_Click()
SQL = "Delete * from Tab_Trans_Aux_Cartao"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

var_Resposta = UCase(InputBox("Digite o número do cartão.", "SisTrans"))
var_SQL = "select * FROM cons_cartao"
var_SQL = var_SQL + " where nr_cartao='" & var_Resposta & "'"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
   MsgBox "Não há registros a listar", vbInformation + vbOKOnly, "SisTrans"
   Exit Sub
Else
   DS.MoveFirst
   Do While Not DS.EOF
      SQL = "select * from Tab_Trans_Aux_Cartao"
      Set ds1 = New Recordset
      ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      ds1.AddNew
      ds1!nr_cartao = DS(0)
      ds1!Ano = DS(1)
      ds1!TARJA = DS(2)
      ds1!Veiculo = DS(3)
      ds1!DT_EMISSAO = DS(4)
      ds1!COD_OM = DS(5)
      ds1!CPF_pessoa = DS(6)
      ds1!cnpj_fornecedor = DS(7)
      ds1!autorizacao = DS(8)
      ds1!data_validade = DS(9)
      ds1!Extraviado = DS(10)
      ds1!Motivo = DS(11)
      ds1!NOME = DS(12)
      ds1.UpdateBatch adAffectAll
      DS.MoveNext
   Loop
   CD.Flags = &H40
   CD.ShowPrinter
   CrystalReport.DataFiles(0) = (xpath)
   CrystalReport.ReportFileName = App.Path + "\cartao.rpt"
   CrystalReport.Action = 1
End If
End Sub
