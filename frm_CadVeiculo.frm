VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_CadVeiculo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SisTrans - Cadastro de Veículos"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6435
   ControlBox      =   0   'False
   Icon            =   "frm_CadVeiculo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Btn_Cartao 
      Caption         =   "&Cartão"
      Enabled         =   0   'False
      Height          =   855
      Left            =   3480
      Picture         =   "frm_CadVeiculo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3960
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Height          =   135
      Left            =   120
      TabIndex        =   27
      Top             =   3720
      Width           =   6135
   End
   Begin VB.CommandButton btn_Salvar 
      Caption         =   "Sal&var"
      Height          =   855
      Left            =   120
      Picture         =   "frm_CadVeiculo.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salvar"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton btn_Sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5280
      Picture         =   "frm_CadVeiculo.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton btn_Excluir 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   855
      Left            =   1080
      Picture         =   "frm_CadVeiculo.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Excluir"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton btn_Editar 
      Caption         =   "E&ditar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   2040
      Picture         =   "frm_CadVeiculo.frx":1412
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Editar"
      Top             =   3960
      Width           =   855
   End
   Begin VB.Frame fra_Cadastro 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6375
      Begin MSMask.MaskEdBox txt_Placa 
         Height          =   300
         Left            =   240
         TabIndex        =   0
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         AutoTab         =   -1  'True
         MaxLength       =   8
         Mask            =   "AAA-9999"
         PromptChar      =   "_"
      End
      Begin VB.Frame fra_Responsavel 
         Caption         =   "Responsável pelo veículo:"
         Height          =   1335
         Left            =   120
         TabIndex        =   22
         Top             =   2280
         Width           =   6135
         Begin VB.OptionButton opt_Identificado 
            Caption         =   "Não identificado"
            Height          =   255
            Left            =   4320
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin MSMask.MaskEdBox txt_Responsavel 
            Height          =   300
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777152
            PromptChar      =   "_"
         End
         Begin VB.Frame Frame5 
            Height          =   135
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   5895
         End
         Begin VB.TextBox txt_Descricao 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   960
            Width           =   3855
         End
         Begin VB.OptionButton opt_OM 
            Caption         =   "OM"
            Height          =   255
            Left            =   3240
            TabIndex        =   8
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton opt_Fornecedor 
            Caption         =   "Fornecedor"
            Height          =   255
            Left            =   1680
            TabIndex        =   7
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton opt_Pessoa 
            Caption         =   "Pessoa"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lbl_Responsavel 
            AutoSize        =   -1  'True
            Caption         =   "Descrição do Responsavel:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   1965
         End
      End
      Begin VB.Frame Fra_Emplacamento 
         Caption         =   "Emplacamento:"
         Height          =   855
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   4695
         Begin VB.ComboBox txt_EmplacUF 
            Height          =   315
            ItemData        =   "frm_CadVeiculo.frx":1854
            Left            =   3960
            List            =   "frm_CadVeiculo.frx":18AC
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txt_EmplacCidade 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Left            =   3960
            TabIndex        =   20
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Fra_Veiculo 
         Caption         =   "Características do veículo:"
         Height          =   975
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   6135
         Begin VB.TextBox txt_Ano 
            Height          =   285
            Left            =   3840
            MaxLength       =   4
            TabIndex        =   4
            Top             =   480
            Width           =   615
         End
         Begin VB.ComboBox Combo_Modelo_Marca_Tipo 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   480
            Width           =   3615
         End
         Begin VB.TextBox txt_Cor 
            Height          =   285
            Left            =   4560
            TabIndex        =   5
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Modelo:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Ano:"
            Height          =   195
            Left            =   3840
            TabIndex        =   24
            Top             =   240
            Width           =   330
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cor:"
            Height          =   195
            Left            =   4560
            TabIndex        =   23
            Top             =   240
            Width           =   285
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Placa:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   450
      End
   End
End
Attribute VB_Name = "frm_CadVeiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vgl_Responsavel As Byte
Dim rs_tabVeiculo As Recordset
Dim rs_Combo As Recordset
Dim rs_Responsavel As Recordset
Dim rs_Cartao As Recordset
Private Sub cmdSair_Click()
Unload Me
End Sub
Private Sub btn_Cartao_Click()

If txt_Placa.Text = "___-____" Then Exit Sub

vgl_placa = txt_Placa
'Envia os dados do form atual para o de Cartao p/
'ser feitas as deidas operacoes

SQL = "select cns_Trans_veiculo_modmarc.*, cns_Trans_veiculo_pessoa.* "
SQL = SQL + "FROM cns_Trans_veiculo_modmarc, cns_Trans_veiculo_pessoa "
SQL = SQL + "where cns_Trans_veiculo_modmarc.PLACA = cns_Trans_veiculo_pessoa.PLACA "
SQL = SQL + "and cns_Trans_veiculo_modmarc.placa = '" & txt_Placa & "'"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount <> 0 Then
   'Trata a tabela de cartao
   vgl_TipoResponsavel = "CPF"
   vgl_PostoGrad = DS!POST_GRAD_CATFUNC
   
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
      frm_Cartao_Emitir.txt_Codigo = txt_Responsavel.Text
      frm_Cartao_Emitir.txt_Descricao = txt_Descricao.Text
      frm_Cartao_Emitir.Show
   End If
Else
   SQL = "select cns_Trans_veiculo_modmarc.*, cns_Trans_veiculo_fornec.* "
   SQL = SQL + "FROM cns_Trans_veiculo_modmarc, cns_Trans_veiculo_fornec "
   SQL = SQL + "where cns_Trans_veiculo_modmarc.PLACA = cns_Trans_veiculo_fornec.PLACA "
   SQL = SQL + "and cns_Trans_veiculo_modmarc.placa = '" & txt_Placa & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   If DS.RecordCount <> 0 Then
      vgl_TipoResponsavel = "CNPJ"
      frm_Cartao_Emitir.txt_Codigo = txt_Responsavel.Text
      frm_Cartao_Emitir.txt_Descricao = txt_Descricao.Text
      vgl_placa = txt_Placa.Text
      frm_Cartao_Emitir.Show
   Else
      SQL = "select cns_Trans_veiculo_modmarc.*, cns_Trans_veiculo_om.* "
      SQL = SQL + "FROM cns_Trans_veiculo_modmarc, cns_Trans_veiculo_om "
      SQL = SQL + "where cns_Trans_veiculo_modmarc.PLACA = cns_Trans_veiculo_om.PLACA "
      SQL = SQL + "and cns_Trans_veiculo_modmarc.placa = '" & txt_Placa & "'"
      Set DS = New Recordset
      DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      If DS.RecordCount <> 0 Then
         vgl_TipoResponsavel = "OM"
         frm_Cartao_Emitir.txt_Codigo = txt_Responsavel
         frm_Cartao_Emitir.txt_Descricao = txt_Descricao.Text
         vgl_placa = txt_Placa.Text
         frm_Cartao_Emitir.Show
      Else
         SQL = "select cns_Trans_veiculo_modmarc.*, cns_Trans_veiculo_null.* "
         SQL = SQL + "FROM cns_Trans_veiculo_modmarc, cns_Trans_veiculo_null "
         SQL = SQL + "where cns_Trans_veiculo_modmarc.PLACA = cns_Trans_veiculo_null.PLACA "
         SQL = SQL + "and cns_Trans_veiculo_modmarc.placa = '" & txt_Placa & "'"
         Set DS = New Recordset
         DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
         If DS.RecordCount <> 0 Then
            MsgBox "Motorista não identificado não se emite cartão.", vbOKOnly + vbInformation, "SisTrans"
         End If
      End If
   End If
End If
End Sub
Private Sub btn_Editar_Click()
VarAcao = "Registro Velho"

txt_Placa.Enabled = False
fra_Cadastro.Enabled = True
btn_Salvar.Enabled = True
btn_Editar.Enabled = False
btn_Excluir.Enabled = False
btn_Cartao.Enabled = False
fra_Cadastro.Enabled = True
fra_Responsavel.Enabled = True
fra_Veiculo.Enabled = True
Fra_Emplacamento.Enabled = True

End Sub
Private Sub btn_Excluir_Click()
On Error GoTo Error

If txt_Placa.Text <> "___-____" Then
   If MsgBox("Deseja excluir o registro?", vbYesNo + vbQuestion, "SisTrans") = vbYes Then
      SQL = "delete * from tab_Trans_Veiculo where placa ='" & txt_Placa.Text & "'"
      Set DS = New Recordset
      DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      MsgBox "Registo Deletado.", vbOKOnly + vbInformation, "SisTrans"
   Else
      MsgBox "Operação cancelada.", vbOKOnly + vbInformation, "SisTrans"
      txt_Placa.SetFocus
      Exit Sub
   End If
   Call LIMPAR
   txt_Placa.Enabled = True
   txt_Placa.Mask = "        "
   txt_Placa.Mask = "AAA-9999"
   txt_Responsavel.Mask = "              "
   txt_Responsavel.Mask = "999.999.999-99"
   txt_Descricao.Text = ""
End If
Exit Sub
Error:
    MsgBox "Não é possível excluir o registro, ele pode fazer parte de um ou mais relacionamentos.Ex.: multa x veículo, veículo x proprietário. Caso seja realmente necessária a exclusão, contate o administrador do Banco de Dados.", vbOKOnly + vbInformation, "SisTrans"
    btn_Sair.SetFocus
End Sub
Private Sub btn_Sair_Click()
Unload Me
End Sub
Private Sub btn_Salvar_Click()
If txt_Placa.Text = "___-____" Then
   MsgBox "Digite a placa.", vbInformation, "SISTRANS"
   Exit Sub
End If

If Combo_Modelo_Marca_Tipo.ListIndex = -1 Then
    MsgBox "É necessário a entrada de um modelo para este veículo.", vbCritical + vbOKOnly, "SisTrans"
    Combo_Modelo_Marca_Tipo.SetFocus
    Exit Sub
End If
  
Call Salvar
Call Form_Load

End Sub
Private Sub Salvar()

On Error GoTo Error
If VarAcao = "Registro Novo" Then
   '-----------------------------------------------------
   'Controla o número máximo de veículos por responsável.
   '-----------------------------------------------------
   SQL = "select * from tab_trans_veiculo"
   SQL = SQL + " where cpf_resp_pessoa = '" & txt_Responsavel.Text & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   If DS.RecordCount > 3 Then
      MsgBox "Este responsavel já possui 4 veículos cadastrados. Exclua algum veículo para cadastrar o desejado.", vbInformation + vbOKOnly, "SisTrans"
      Exit Sub
   End If
  '-----------------------------------------------
End If

SQL = "select * from tab_Trans_Veiculo"
SQL = SQL + " where placa ='" & txt_Placa.Text & "'"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
With DS
     If .RecordCount = 0 Then
        .AddNew
        !placa = txt_Placa.Text
     End If
    
    !cidade_empl = txt_EmplacCidade.Text
    !uf_empl = txt_EmplacUF.Text
    
    If Combo_Modelo_Marca_Tipo.ListIndex <> -1 Then
        !Cod_Modelo = Combo_Modelo_Marca_Tipo.ItemData(frm_CadVeiculo.Combo_Modelo_Marca_Tipo.ListIndex)
    End If
    
    !Cor_Pred = txt_Cor.Text
    !Ano = txt_Ano.Text
    If opt_Pessoa.Value = True Then
       !cpf_resp_pessoa = txt_Responsavel
       !cnpj = Null
       !COD_OM = Null
    ElseIf opt_Fornecedor.Value = True Then
       !cnpj = txt_Responsavel
       !cpf_resp_pessoa = Null
       !COD_OM = Null
    ElseIf opt_OM.Value = True Then
       !COD_OM = txt_Responsavel
       !cnpj = Null
       !cpf_resp_pessoa = Null
    ElseIf opt_Pessoa.Value = False And opt_Fornecedor.Value = False And opt_OM.Value = False Then
       !COD_OM = Null
       !cnpj = Null
       !cpf_resp_pessoa = Null
    End If
    .UpdateBatch adAffectAll
    
    If opt_Identificado.Value = False Then
        If MsgBox("Arquivo salvo. Deseja emitir cartão?", vbYesNo + vbQuestion, "SisTrans") = vbYes Then
             Call btn_Cartao_Click
        End If
    Else
        MsgBox "Arquivo salvo", vbOKOnly + vbInformation, "SisTrans"
    End If
    
End With

btn_Salvar.Enabled = False
btn_Editar.Enabled = False
btn_Excluir.Enabled = False
btn_Cartao.Enabled = False
Exit Sub
Error:
    MsgBox "Erro. Responsável não identificado.", vbOKOnly + vbInformation, "SisTrans"
    btn_Sair.SetFocus
End Sub
Private Sub Form_Activate()
If vgl_Responsavel = 4 Then
    opt_Identificado.Value = True
    btn_Cartao.Enabled = False
End If
End Sub
Private Sub Form_Load()
vgl_Cor = ""
var_Resposta = ""
Me.Top = 0
Me.Left = 0

Call LIMPAR

txt_Placa.Mask = "        "
txt_Placa.Mask = "AAA-9999"
txt_EmplacUF.ListIndex = 0

SQL = "SELECT Tab_Trans_Modelo_Veic.Modelo, Tab_Trans_Marca_Veic.Marca,"
SQL = SQL + " Tab_Trans_Modelo_Veic.Tipo, Tab_Trans_Modelo_Veic.Cod"
SQL = SQL + " FROM Tab_Trans_Marca_Veic, Tab_Trans_Modelo_Veic "
SQL = SQL + " where Tab_Trans_Marca_Veic.Cod = Tab_Trans_Modelo_Veic.Cod_Marca"
SQL = SQL + " ORDER BY Tab_Trans_Modelo_Veic.Tipo, Tab_Trans_Marca_Veic.Marca, Tab_Trans_Modelo_Veic.Modelo;"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

Combo_Modelo_Marca_Tipo.Clear
Do While Not DS.EOF
   Combo_Modelo_Marca_Tipo.AddItem ((DS(0)) + "-" + DS(1) + "-" + DS(2))
   Combo_Modelo_Marca_Tipo.ItemData(Combo_Modelo_Marca_Tipo.NewIndex) = DS(3)
   DS.MoveNext
Loop
If vgl_Responsavel = 4 Then
    opt_OM.Value = True
    opt_Identificado.Value = True
    btn_Cartao.Enabled = False
End If
End Sub
Private Sub opt_Fornecedor_Click()
txt_Descricao.Enabled = True
txt_Responsavel.Enabled = True
lbl_Responsavel.Caption = "Entre com o CNPJ:"
txt_Responsavel.Mask = "                "
txt_Descricao.Text = " "
txt_Responsavel.Mask = "99999999/9999-99"
If txt_Responsavel.Text <> "________/____-__" Then
   Call txt_Responsavel_LostFocus
End If
btn_Cartao.Enabled = True
End Sub
Private Sub opt_Identificado_Click()
txt_Descricao.Text = "Responsável não identificado"
txt_Responsavel.Mask = "              "
txt_Descricao.Enabled = False
txt_Responsavel.Enabled = False
btn_Cartao.Enabled = False
End Sub


Private Sub opt_OM_Click()
txt_Descricao.Enabled = True
txt_Responsavel.Enabled = True
lbl_Responsavel.Caption = "Entre com o Código:"
txt_Responsavel.Mask = "              "
txt_Descricao.Text = " "
txt_Responsavel.Mask = "99999"
If txt_Responsavel.Text <> "_____" Then
   Call txt_Responsavel_LostFocus
End If
btn_Cartao.Enabled = True
End Sub
Private Sub opt_Pessoa_Click()
txt_Descricao.Enabled = True
txt_Responsavel.Enabled = True
lbl_Responsavel.Caption = "Entre com o CPF:"
txt_Responsavel.Mask = "              "
txt_Descricao.Text = " "
txt_Responsavel.Mask = "999.999.999-99"
If txt_Responsavel.Text <> "___.___.___-__" Then
   Call txt_Responsavel_LostFocus
End If
btn_Cartao.Enabled = True
End Sub
Private Sub txt_Cor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txt_EmplacCidade_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txt_Placa_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txt_Placa_LostFocus()
If txt_Placa.Text = "___-____" Then Exit Sub
vgl_placa = txt_Placa
SQL = "select * "
SQL = SQL + " from cns_trans_veiculo"
SQL = SQL + " where placa ='" & txt_Placa.Text & "'"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
   SQL = "select * "
   SQL = SQL + " from cns_trans_veiculo_null"
   SQL = SQL + " where placa ='" & txt_Placa.Text & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   With DS
        If .RecordCount <> 0 Then
           var_pessoa = !cpf_resp_pessoa
           Var_Empresa = !cnpj
           var_OM = !COD_OM
     
           
           VarAcao = "Registro Velho"
           txt_EmplacCidade.Text = !cidade_empl
                
           'Posiciona a combo UF no item do arquivo.
   
           VarIndex = 0
           txt_EmplacUF.ListIndex = 0
           Do While Not (txt_EmplacUF.Text = DS(2))
              If VarIndex <= (txt_EmplacUF.ListCount - 1) Then
                 VarIndex = (txt_EmplacUF.ListIndex + 1)
                 txt_EmplacUF.ListIndex = VarIndex
              End If
           Loop
           txt_EmplacUF.ListIndex = VarIndex
                
           txt_Cor.Text = !Cor_Pred
           txt_Ano.Text = !Ano
           'Posiciona a combo no item do arquivo.

           VarIndex = 0
           frm_CadVeiculo.Combo_Modelo_Marca_Tipo.ListIndex = 0
           Do While Not (frm_CadVeiculo.Combo_Modelo_Marca_Tipo.ItemData(frm_CadVeiculo.Combo_Modelo_Marca_Tipo.ListIndex)) = DS(8)
              If VarIndex <= (frm_CadVeiculo.Combo_Modelo_Marca_Tipo.ListCount - 1) Then
                 VarIndex = (frm_CadVeiculo.Combo_Modelo_Marca_Tipo.ListIndex + 1)
                 frm_CadVeiculo.Combo_Modelo_Marca_Tipo.ListIndex = VarIndex
              End If
           Loop
           frm_CadVeiculo.Combo_Modelo_Marca_Tipo.ListIndex = VarIndex

           If !cpf_resp_pessoa <> " " Then
              opt_Pessoa.Value = True
              txt_Responsavel.Mask = "              "
              txt_Responsavel.Mask = "999.999.999-99"
              txt_Responsavel = !cpf_resp_pessoa
              txt_Descricao.Text = !NOME
               btn_Cartao.Enabled = True
           ElseIf !cnpj <> " " Then
              opt_Fornecedor.Value = True
              txt_Responsavel.Mask = "                 "
              txt_Responsavel.Mask = "99999999/9999-99"
              txt_Responsavel = !cnpj
              txt_Descricao.Text = !NOME
               btn_Cartao.Enabled = True
           ElseIf !COD_OM <> " " Then
              opt_OM.Value = True
              txt_Responsavel.Mask = "     "
              txt_Responsavel.Mask = "99999"
              txt_Responsavel = !COD_OM
              txt_Descricao.Text = !NOME
               btn_Cartao.Enabled = True
           Else
              opt_Identificado.Value = True
               btn_Cartao.Enabled = False
           End If
           btn_Editar.Enabled = True
           btn_Excluir.Enabled = True
           btn_Salvar.Enabled = False
           fra_Responsavel.Enabled = False
           fra_Veiculo.Enabled = False
           Fra_Emplacamento.Enabled = False
       Else
           VarAcao = "Registro Novo"
           Call LIMPAR
           fra_Cadastro.Enabled = True
           btn_Salvar.Enabled = True
           btn_Editar.Enabled = False
           btn_Excluir.Enabled = False
           btn_Cartao.Enabled = False
           fra_Responsavel.Enabled = True
           fra_Veiculo.Enabled = True
           Fra_Emplacamento.Enabled = True
       End If
    End With
Else
    With DS
         VarAcao = "Registro Velho"
         txt_EmplacCidade.Text = !cidade_empl
               
         'Posiciona a combo UF no item do arquivo.
   
         VarIndex = 0
         txt_EmplacUF.ListIndex = 0
         Do While Not (txt_EmplacUF.Text = DS(2))
            If VarIndex <= (txt_EmplacUF.ListCount - 1) Then
               VarIndex = (txt_EmplacUF.ListIndex + 1)
               txt_EmplacUF.ListIndex = VarIndex
            End If
         Loop
         txt_EmplacUF.ListIndex = VarIndex
                
         txt_Cor.Text = !Cor_Pred
         txt_Ano.Text = !Ano
         'Posiciona a combo no item do arquivo.

         VarIndex = 0
         frm_CadVeiculo.Combo_Modelo_Marca_Tipo.ListIndex = 0
         Do While Not (frm_CadVeiculo.Combo_Modelo_Marca_Tipo.ItemData(frm_CadVeiculo.Combo_Modelo_Marca_Tipo.ListIndex)) = DS(8)
            If VarIndex <= (frm_CadVeiculo.Combo_Modelo_Marca_Tipo.ListCount - 1) Then
               VarIndex = (frm_CadVeiculo.Combo_Modelo_Marca_Tipo.ListIndex + 1)
               frm_CadVeiculo.Combo_Modelo_Marca_Tipo.ListIndex = VarIndex
            End If
         Loop
         frm_CadVeiculo.Combo_Modelo_Marca_Tipo.ListIndex = VarIndex

         If !cpf_resp_pessoa <> " " Then
            opt_Pessoa.Value = True
            txt_Responsavel.Mask = "              "
            txt_Responsavel.Mask = "999.999.999-99"
            txt_Responsavel = !cpf_resp_pessoa
            txt_Descricao.Text = !NOME
            btn_Cartao.Enabled = True
         ElseIf !cnpj <> " " Then
            opt_Fornecedor.Value = True
            txt_Responsavel.Mask = "                 "
            txt_Responsavel.Mask = "99999999/9999-99"
            txt_Responsavel = !cnpj
            txt_Descricao.Text = !NOME
            btn_Cartao.Enabled = True
         ElseIf !COD_OM <> " " Then
            opt_OM.Value = True
            txt_Responsavel.Mask = "     "
            txt_Responsavel.Mask = "99999"
            txt_Responsavel = !COD_OM
            txt_Descricao.Text = !NOME
            btn_Cartao.Enabled = True
         Else
            opt_Identificado.Value = True
            btn_Cartao.Enabled = False
         End If
     
         btn_Editar.Enabled = True
         btn_Excluir.Enabled = True
         btn_Salvar.Enabled = False
         fra_Responsavel.Enabled = False
         fra_Veiculo.Enabled = False
         Fra_Emplacamento.Enabled = False
    End With
End If
End Sub
Private Sub LIMPAR()
txt_EmplacCidade.Text = ""
txt_EmplacUF.ListIndex = 0
Combo_Modelo_Marca_Tipo.ListIndex = -1
txt_Cor.Text = ""
txt_Ano.Text = ""
End Sub
Private Sub txt_Responsavel_LostFocus()
If txt_Responsavel.Text = "" Then Exit Sub
If opt_Pessoa.Value = True Then
   SQL = "select * from tab_ger_pessoa"
   SQL = SQL + " where cpf ='" & txt_Responsavel.Text & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   If DS.RecordCount <> 0 Then
      txt_Descricao.Text = DS!NOME
   Else
      txt_Descricao.Text = ""
      frm_CadPessoa.TXT_CPF.Mask = "              "
      frm_CadPessoa.TXT_CPF.Mask = "999.999.999-99"
      frm_CadPessoa.TXT_CPF = txt_Responsavel
   End If
ElseIf opt_Fornecedor.Value = True Then
       SQL = "select * from tab_trans_fornecedor"
       SQL = SQL + " where cnpj ='" & txt_Responsavel.Text & "'"
       Set DS = New Recordset
       DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
       If DS.RecordCount <> 0 Then
          txt_Descricao.Text = DS!NOME
       Else
          txt_Descricao.Text = ""
          frm_CadFornecedor.txt_CNPJ.Mask = "                "
          frm_CadFornecedor.txt_CNPJ.Mask = "99999999/9999-99"
          frm_CadFornecedor.txt_CNPJ.Text = txt_Responsavel
       End If
ElseIf opt_OM.Value = True Then
       SQL = "select * from tab_ger_om"
       SQL = SQL + " where COD ='" & txt_Responsavel.Text & "'"
       Set DS = New Recordset
       DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
       If DS.RecordCount <> 0 Then
          txt_Descricao.Text = DS!NOME
       Else
          txt_Descricao.Text = ""
          Frm_OM.TXT_COD_OM.Mask = "     "
          Frm_OM.TXT_COD_OM.Mask = "99999"
          Frm_OM.TXT_COD_OM.Text = txt_Responsavel
       End If
End If
End Sub
