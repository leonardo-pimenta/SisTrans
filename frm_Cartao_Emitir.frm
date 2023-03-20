VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_Cartao_Emitir 
   BackColor       =   &H80000013&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SisTrans - Emissão de Cartão"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3885
   Icon            =   "frm_Cartao_Emitir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport CrystalReport 
      Left            =   2520
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "L:\Sistrans\Fonte atualizada 09MAR2004\Fonte\Pessoa.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TXT_Motivo 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   27
      Top             =   3960
      Width           =   3255
   End
   Begin VB.CheckBox Extraviado 
      Caption         =   "Extraviado"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   3720
      Width           =   1095
   End
   Begin VB.ComboBox CMB_VEICULO 
      Height          =   315
      ItemData        =   "frm_Cartao_Emitir.frx":0442
      Left            =   1920
      List            =   "frm_Cartao_Emitir.frx":044C
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txt_Autorizado 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Txt_Validade 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Txt_Tarja 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txt_NumCartao 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   240
      MaxLength       =   10
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txt_Ano 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton BTN_SALVAR 
      Caption         =   "Sal&var"
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      Picture         =   "frm_Cartao_Emitir.frx":0468
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salvar"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton btn_Editar 
      Caption         =   "E&ditar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   1080
      Picture         =   "frm_Cartao_Emitir.frx":08AA
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Editar"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton btn_Excluir 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   855
      Left            =   2040
      Picture         =   "frm_Cartao_Emitir.frx":0CEC
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Excluir"
      Top             =   4320
      Width           =   855
   End
   Begin VB.Frame FRA_CARTAO 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   3495
      Begin VB.Frame Frame2 
         Caption         =   "Responsável:"
         Height          =   1215
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   3255
         Begin VB.TextBox txt_Codigo 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            TabIndex        =   8
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txt_Descricao 
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            TabIndex        =   9
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lbl_Codigo 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   765
         End
      End
      Begin VB.TextBox Txt_DtEmissao 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Txt_TotalEmitidos 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Data de Emissão:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label Label8 
         Caption         =   "Total de Cartões:"
         Height          =   195
         Left            =   1800
         TabIndex        =   18
         Top             =   0
         Width           =   1275
      End
   End
   Begin VB.CommandButton btn_Cartao 
      Caption         =   "&Emitir"
      Enabled         =   0   'False
      Height          =   855
      Left            =   3000
      Picture         =   "frm_Cartao_Emitir.frx":112E
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Editar"
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Autorizado por:"
      Height          =   195
      Left            =   1920
      TabIndex        =   25
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label Label9 
      Caption         =   "Validade:"
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tarja:"
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Nº Cartão:"
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ano:"
      Height          =   195
      Left            =   1920
      TabIndex        =   21
      Top             =   0
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   195
      Left            =   1920
      TabIndex        =   20
      Top             =   600
      Width           =   360
   End
End
Attribute VB_Name = "frm_Cartao_Emitir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_Cartao As Recordset
Dim rs_Cartao_CNPJ As Recordset
Dim rs_Cartao_OM As Recordset
Dim rs_AuxCartao As Recordset
Dim var_Aux1 As String
Dim var_Aux2 As String
Dim var_Aux3 As String
Dim var_Controle As Integer
Dim var_NumVeiculo As Byte
'padrao para todos os responsaveis
Private Sub btn_Sair_Click()
Unload Me
End Sub
Private Sub btn_Cartao_Click()
If txt_NumCartao.Text = "" Then Exit Sub
If txt_Ano.Text = "" Then Exit Sub

var_Aux1 = ""
var_Aux2 = ""
var_Aux3 = ""
vgl_TipoResponsavell = ""
vgl_TipoResponsavel = ""
vgl_TipoResponsavelll = ""

'--------------------------------------------------
'Apaga todos os dados decontrole na tabela auxiliar
'--------------------------------------------------
 Set rs_AuxCartao = New Recordset
 rs_AuxCartao.Open "DELETE * FROM tab_trans_aux_cartao_emitir", cnConexao, adOpenStatic, adLockOptimistic
'--------------------------------------------------

var_SQL = "select * from cons_cartao"
var_SQL = var_SQL + " where cpf_pessoa='" & txt_Codigo & "'"
var_SQL = var_SQL + " and  nr_cartao='" & txt_NumCartao & "'"
var_SQL = var_SQL + " AND Tarja = '" & Txt_Tarja.Text & "'"
var_SQL = var_SQL + " AND ano = '" & txt_Ano.Text & "'"
var_SQL = var_SQL + " AND veiculo = '" & CMB_VEICULO.Text & "'"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount <> 0 Then
   If DS!data_validade <> "" Then
      vgl_TipoResponsavell = "Provisorio"
   End If
   vgl_TipoResponsavel = "CPF"
Else
   var_SQL = "select * from cons_cartao"
   var_SQL = var_SQL + " where cnpj_fornecedor='" & txt_Codigo & "'"
   var_SQL = var_SQL + " and  nr_cartao='" & txt_NumCartao & "'"
   var_SQL = var_SQL + " AND Tarja = '" & Txt_Tarja.Text & "'"
   var_SQL = var_SQL + " AND ano = '" & txt_Ano.Text & "'"
   var_SQL = var_SQL + " AND veiculo = '" & CMB_VEICULO.Text & "'"
   Set DS = New Recordset
   DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
   If DS.RecordCount <> 0 Then
      If DS!data_validade <> "" Then
         vgl_TipoResponsavell = "Provisorio"
      End If
      vgl_TipoResponsavel = "CNPJ"
   Else
      var_SQL = "select * from cons_cartao"
      var_SQL = var_SQL + " where cod_om='" & txt_Codigo & "'"
      var_SQL = var_SQL + " and  nr_cartao='" & txt_NumCartao & "'"
      var_SQL = var_SQL + " AND Tarja = '" & Txt_Tarja.Text & "'"
      var_SQL = var_SQL + " AND ano = '" & txt_Ano.Text & "'"
      var_SQL = var_SQL + " AND veiculo = '" & CMB_VEICULO.Text & "'"
      Set DS = New Recordset
      DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
      If DS.RecordCount <> 0 Then
         If DS!data_validade <> "" Then
            vgl_TipoResponsavell = "Provisorio"
         End If
         If DS!TARJA = "AZUL" Then
            vgl_TipoResponsavelll = ""
         Else
            vgl_TipoResponsavelll = "OFICIAL"
         End If
         vgl_TipoResponsavel = "OM"
      Else
         Exit Sub
      End If
   End If
End If

If vgl_TipoResponsavel = "CPF" Then
    Set rs_Cartao = New Recordset
    With rs_Cartao
        
        var_SQL = "select cns_Trans_veiculo_pessoa.*, cns_Trans_veiculo.* "
        var_SQL = var_SQL + "FROM cns_Trans_veiculo_pessoa, cns_Trans_veiculo "
        var_SQL = var_SQL + "where cns_Trans_veiculo_pessoa.PLACA = cns_Trans_veiculo.PLACA "
        var_SQL = var_SQL + "and cns_Trans_veiculo_pessoa.CPF = '" & txt_Codigo & "'"
        var_SQL = var_SQL + "and cns_Trans_veiculo_pessoa.Tipo = '" & CMB_VEICULO & "'"

        .Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
        If .RecordCount = 0 Then
            MsgBox "Esta PESSOA não possui " + "'" & CMB_VEICULO & "'", vbInformation + vbOKOnly, "SisTrans"
            .Close
        Else
            .MoveFirst
            For x = 1 To .RecordCount
                var_Aux1 = !placa + " / " + var_Aux1
                var_Aux2 = !modelo + " / " + var_Aux2
                
                .MoveNext
            Next
            
            .MoveFirst
            vgl_PostoGrad = rs_Cartao!POST_GRAD_CATFUNC
            Set rs_AuxCartao = New Recordset
            With rs_AuxCartao
                .Open "select * from tab_trans_aux_cartao_emitir", cnConexao, adOpenStatic, adLockOptimistic
                .AddNew
                    !nr_cartao = txt_NumCartao.Text
                    !aux1 = var_Aux1
                    !aux2 = var_Aux2
                    !aux3 = rs_Cartao!NOME
                    !aux4 = vgl_PostoGrad
                    !aux5 = rs_Cartao!SIGLA
                    !aux6 = Txt_Tarja
                .UpdateBatch adAffectAll
                .Close
            End With
        
        End If
    End With

ElseIf vgl_TipoResponsavel = "CNPJ" Then

    Set rs_Cartao_CNPJ = New Recordset
    With rs_Cartao_CNPJ
    
        var_SQL = "select * FROM TAB_Trans_veiculo,tab_trans_modelo_veic "
        var_SQL = var_SQL + " where tab_trans_modelo_veic.cod=tab_trans_veiculo.cod_modelo"
        var_SQL = var_SQL + " and tab_Trans_modelo_veic.Tipo = '" & CMB_VEICULO & "'"
        var_SQL = var_SQL + " and cnpj = '" & txt_Codigo & "'"
        
        .Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
        
        'lança o num de veiculos para salvar depois na tabela de cartao
        'var_NumVeiculo = .RecordCount
        
        If .RecordCount = 0 Then
            MsgBox "Este fornecedor não possui veículos.", vbOKOnly + vbInformation, "SisTrans"
            Exit Sub
        Else
            .MoveFirst
        End If
        
        For X1 = 1 To .RecordCount
        
            var_Aux1 = ""
            var_Aux2 = ""
            var_Aux3 = ""

'''''
            Set rs_Cartao = New Recordset
            With rs_Cartao
            
                var_SQL = "select * FROM cns_Trans_veiculo_FORNEC "
                var_SQL = var_SQL + "where cnpj = '" & txt_Codigo & "'"
                var_SQL = var_SQL + "and PLACA = '" & rs_Cartao_CNPJ!placa & "'"
                
                .Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
                
                If .RecordCount = 0 Then
                    MsgBox "Este fornecedor não possui motoristas.", vbInformation + vbOKOnly, "SisTrans"
                    .Close
                Else
                    .MoveFirst
                    For X2 = 1 To .RecordCount
                        var_Aux1 = !nome_motorista + " / " + var_Aux1
                        .MoveNext
                    Next
                    
                    .MoveFirst
                    
                    Set rs_AuxCartao = New Recordset
                    With rs_AuxCartao
                        .Open "select * from tab_trans_aux_cartao_emitir", cnConexao, adOpenStatic, adLockOptimistic
                        .AddNew
                            !nr_cartao = txt_NumCartao
                            !aux1 = rs_Cartao!placa
                            !aux2 = rs_Cartao!modelo
                            !aux3 = txt_Codigo.Text
                            !aux4 = txt_Descricao.Text
                            !aux5 = var_Aux1
                            !aux6 = Txt_Tarja
                        .UpdateBatch adAffectAll
                        .Close
                    End With
                    
                End If
            End With
            .MoveNext
        Next

    End With
    
ElseIf vgl_TipoResponsavel = "OM" And vgl_TipoResponsavelll <> "OFICIAL" Then
       var_SQL = "select * FROM TAB_Trans_veiculo,tab_trans_mod_veic "
       var_SQL = var_SQL + "where cod_om = '" & txt_Codigo & "'"
       var_SQL = var_SQL + " and tab_trans_mod_veic.cod=tab_trans_veiculo.cod_modelo"
       var_SQL = var_SQL + " and tab_Trans_mod_veic.Tipo = '" & CMB_VEICULO & "'"
       
       Set DS = New Recordset
       DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
          
       SQL = "select * from tab_trans_aux_cartao_emitir"
       Set ds1 = New Recordset
       ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
          
       ds1.AddNew
       ds1!nr_cartao = txt_NumCartao.Text
       ds1!aux1 = "----//----"
       ds1!aux2 = "----//----"
       ds1!aux3 = txt_Codigo.Text
       ds1!aux4 = txt_Descricao.Text
       ds1!aux5 = ""
       ds1!aux6 = Txt_Tarja
       ds1.UpdateBatch adAffectAll
ElseIf vgl_TipoResponsavel = "OM" And vgl_TipoResponsavelll = "OFICIAL" Then
       var_SQL = "select * FROM TAB_Trans_veiculo,tab_trans_modelo_veic "
       var_SQL = var_SQL + "where cod_om = '" & txt_Codigo & "'"
       var_SQL = var_SQL + " and tab_trans_modelo_veic.cod=tab_trans_veiculo.cod_modelo"
       var_SQL = var_SQL + " and tab_Trans_modelo_veic.Tipo = '" & CMB_VEICULO & "'"
       var_SQL = var_SQL + " and placa='" & vgl_placa & "'"
       Set DS = New Recordset
       DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
       If DS.RecordCount = 0 Then
          Resp = MsgBox("Selecione um veículo", vbCritical, "SisTrans")
          Exit Sub
       End If
       SQL = "select * from tab_trans_aux_cartao_emitir"
       Set ds1 = New Recordset
       ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
          
       ds1.AddNew
       ds1!nr_cartao = txt_NumCartao.Text
       ds1!aux1 = DS!placa
       ds1!aux2 = DS!modelo
       ds1!aux3 = txt_Codigo.Text
       ds1!aux4 = txt_Descricao.Text
       ds1!aux5 = ""
       ds1!aux6 = Txt_Tarja
       ds1.UpdateBatch adAffectAll
End If
'trata qual FORM que solicitou o cartão

If vgl_TipoResponsavel = "CPF" Then
    
    var_SQL = "select * from tab_trans_aux_cartao_emitir "
    var_SQL = var_SQL + "WHERE nr_cartao = '" & txt_NumCartao.Text & "' "
    
    Set rs_AuxCartao = New Recordset
    rs_AuxCartao.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
    
    If vgl_TipoResponsavel = "CPF" And vgl_TipoResponsavell = "" Then
       CD.Flags = &H40
       'CD.ShowPrinter
       CrystalReport.DataFiles(0) = (xpath)
       CrystalReport.ReportFileName = App.Path + "\Pessoa.rpt"
       CrystalReport.Action = 1
    Else
       CD.Flags = &H40
       'CD.ShowPrinter - abre a caixa de configuração da impressora. Sem essa linha,
       'abre a caixa de configuração do Crystall Reports direto.
       CrystalReport.DataFiles(0) = (xpath)
       CrystalReport.ReportFileName = App.Path + "\Provisorio.rpt"
       CrystalReport.Action = 1
    End If
ElseIf vgl_TipoResponsavel = "CNPJ" Then
    
    var_SQL = "select * from tab_trans_aux_cartao_emitir "
    var_SQL = var_SQL + "WHERE AUX3 = '" & txt_Codigo.Text & "' "
    
    If vgl_placa <> "" Then
        var_SQL = "select * from tab_trans_aux_cartao_emitir "
        var_SQL = var_SQL + "WHERE AUX1 = '" & vgl_placa & "' "
    End If
    
    Set rs_AuxCartao = New Recordset
    rs_AuxCartao.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
       
    If vgl_TipoResponsavel = "CNPJ" And vgl_TipoResponsavell = "" Then
       CD.Flags = &H40
       'CD.ShowPrinter
       CrystalReport.DataFiles(0) = (xpath)
       CrystalReport.ReportFileName = App.Path + "\Fornecedor.rpt"
       CrystalReport.Action = 1
    Else
       CD.Flags = &H40
       'CD.ShowPrinter
       CrystalReport.DataFiles(0) = (xpath)
       CrystalReport.ReportFileName = App.Path + "\Provisorio.rpt"
       CrystalReport.Action = 1
    End If
    
ElseIf vgl_TipoResponsavel = "OM" Then

    var_SQL = "select * from tab_trans_aux_cartao_emitir "
    var_SQL = var_SQL + "WHERE AUX3 = '" & txt_Codigo.Text & "' "
    
    If vgl_placa <> "" Then
        var_SQL = "select * from tab_trans_aux_cartao_emitir "
        var_SQL = var_SQL + "WHERE AUX1 = '" & vgl_placa & "' "
    End If
    
    Set rs_AuxCartao = New Recordset
    rs_AuxCartao.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic

    If vgl_TipoResponsavel = "OM" And vgl_TipoResponsavell = "" Then
       CD.Flags = &H40
       'CD.ShowPrinter
       CrystalReport.DataFiles(0) = (xpath)
       CrystalReport.ReportFileName = App.Path + "\OM.rpt"
       CrystalReport.Action = 1
    Else
       CD.Flags = &H40
       'CD.ShowPrinter
       CrystalReport.DataFiles(0) = (xpath)
       CrystalReport.ReportFileName = App.Path + "\Provisorio.rpt"
       CrystalReport.Action = 1
    End If
   
End If

vgl_placa = ""
var_Controle = 0

End Sub

Private Sub btn_Editar_Click()
FRA_CARTAO.Enabled = True
CMB_VEICULO.Enabled = False
txt_NumCartao.Enabled = False
txt_Ano.Enabled = False
Txt_Tarja.Enabled = False
txt_Autorizado.Enabled = True
Txt_Validade.Enabled = True
Extraviado.Enabled = True
If Extraviado.Value = 1 Then
   TXT_Motivo.Enabled = True
End If
btn_Salvar.Enabled = True
btn_Editar.Enabled = False
btn_Excluir.Enabled = False
btn_Cartao.Enabled = False
End Sub
Private Sub btn_Excluir_Click()

var_SQL = "select * FROM tab_Trans_cartao "
var_SQL = var_SQL + "WHERE nr_cartao = '" & txt_NumCartao & "' "
var_SQL = var_SQL + "AND Tarja = '" & Txt_Tarja.Text & "'"
var_SQL = var_SQL + "AND ano = '" & txt_Ano.Text & "'"
var_SQL = var_SQL + "AND veiculo = '" & CMB_VEICULO.Text & "'"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic

If MsgBox("Deseja excluir o registro?", vbYesNo + vbQuestion, "SisTrans") = vbYes Then
   var_SQL = "delete * from tab_Trans_cartao "
   var_SQL = var_SQL + "WHERE nr_cartao = '" & txt_NumCartao & "' "
   var_SQL = var_SQL + "AND Tarja = '" & Txt_Tarja.Text & "'"
   var_SQL = var_SQL + "AND ano = '" & txt_Ano.Text & "'"
   var_SQL = var_SQL + "AND veiculo = '" & CMB_VEICULO.Text & "'"
   Set DS = New Recordset
   DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
   MsgBox "Registo Deletado.", vbOKOnly + vbInformation, "SisTrans"
Else
   MsgBox "Operação cancelada.", vbOKOnly + vbInformation, "SisTrans"
   txt_NumCartao.SetFocus
   Exit Sub
End If
Call LIMPAR
txt_NumCartao = ""
btn_Excluir.Enabled = False
End Sub
Private Sub btn_Salvar_Click()

'IMPORTANTE''''''''''''''''''''''''''''''''''''''''''''''''''''''
var_SQL = "select * FROM tab_Trans_cartao "
var_SQL = var_SQL + " WHERE nr_cartao = '" & txt_NumCartao & "'"
var_SQL = var_SQL + "AND Tarja = '" & Txt_Tarja.Text & "'"
var_SQL = var_SQL + "AND ano = '" & txt_Ano.Text & "'"
var_SQL = var_SQL + "AND veiculo = '" & CMB_VEICULO.Text & "'"

If vgl_TipoResponsavel = "OM" Then
   var_SQL = var_SQL + "and  cod_om = '" & txt_Codigo & "'"
ElseIf vgl_TipoResponsavel = "CNPJ" Then
   var_SQL = var_SQL + "and  cnpj_fornecedor = '" & txt_Codigo & "'"
ElseIf vgl_TipoResponsavel = "CPF" Then
   var_SQL = var_SQL + "and CPF_pessoa = '" & txt_Codigo & "'"
End If
Set DS = New Recordset
With DS
    .Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
    If .RecordCount = 1 Then
       !Ano = txt_Ano.Text
       !data_validade = Txt_Validade
       !autorizacao = txt_Autorizado.Text
       If Extraviado.Value = 1 Then
         !Extraviado = "Sim"
       ElseIf Extraviado.Value = 0 Then
         !Extraviado = ""
       End If
       !Motivo = TXT_Motivo.Text
       .UpdateBatch adAffectAll
        
    ElseIf .RecordCount = 0 Then
           .AddNew
           !nr_cartao = txt_NumCartao.Text
           !Ano = txt_Ano.Text
           !TARJA = Txt_Tarja.Text
           !data_validade = var_Resposta
           !Veiculo = CMB_VEICULO.Text
           !DT_EMISSAO = Txt_DtEmissao.Text
           !autorizacao = txt_Autorizado
           !data_validade = Txt_Validade
         
           If vgl_TipoResponsavel = "CPF" Then
              !CPF_pessoa = txt_Codigo.Text
           ElseIf vgl_TipoResponsavel = "CNPJ" Then
               !cnpj_fornecedor = txt_Codigo.Text
           ElseIf vgl_TipoResponsavel = "OM" Then
               !COD_OM = txt_Codigo.Text
           End If
           .UpdateBatch adAffectAll
    End If
End With
Call LIMPAR
txt_NumCartao = ""
Txt_Tarja.Enabled = True
CMB_VEICULO.Enabled = True
txt_Ano.Enabled = True
txt_NumCartao.Enabled = True
btn_Salvar.Enabled = False
End Sub
Private Sub CMB_VEICULO_Change()
If vgl_TipoResponsavel = "CPF" Then
   lbl_Codigo.Caption = "CPF:"
   'dados para o contador de emissoes
    var_SQL = "select * FROM tab_Trans_cartao "
    var_SQL = var_SQL + "WHERE cpf_pessoa = '" & txt_Codigo.Text & "'"
    Set DS = New Recordset
    DS.Open var_SQL, cnConexao
    Txt_TotalEmitidos.Text = DS.RecordCount
ElseIf vgl_TipoResponsavel = "CNPJ" Then
    lbl_Codigo.Caption = "CNPJ:"
    Txt_Tarja.Text = "NÃO HÁ"
    'dados para o contador de emissoes
    var_SQL = "select * FROM tab_Trans_cartao "
    var_SQL = var_SQL + "WHERE cnpj_fornecedor = '" & txt_Codigo.Text & "'"
    Set DS = New Recordset
    DS.Open var_SQL, cnConexao
    Txt_TotalEmitidos.Text = DS.RecordCount
ElseIf vgl_TipoResponsavel = "OM" Then
    lbl_Codigo.Caption = "Cod. OM:"
    If vgl_Cor = "Azul" Then
       frm_Cartao_Emitir.Txt_Tarja.Text = "AZUL"
    Else
       frm_Cartao_Emitir.Txt_Tarja.Text = "LILÁS"
    End If
    'dados para o contador de emissoes
    var_SQL = "select * FROM tab_Trans_cartao "
    var_SQL = var_SQL + "WHERE cod_om = '" & txt_Codigo.Text & "'"
    Set DS = New Recordset
    DS.Open var_SQL, cnConexao
    Txt_TotalEmitidos.Text = DS.RecordCount
End If
Call Txt_Tarja_LostFocus
End Sub
Private Sub CMB_VEICULO_Click()
If vgl_TipoResponsavel = "CPF" Then
   lbl_Codigo.Caption = "CPF:"
   'dados para o contador de emissoes
    var_SQL = "select * FROM tab_Trans_cartao "
    var_SQL = var_SQL + "WHERE cpf_pessoa = '" & txt_Codigo.Text & "'"
    var_SQL = var_SQL + "and veiculo = '" & CMB_VEICULO.Text & "'"
    Set DS = New Recordset
    DS.Open var_SQL, cnConexao
    Txt_TotalEmitidos.Text = DS.RecordCount
ElseIf vgl_TipoResponsavel = "CNPJ" Then
    lbl_Codigo.Caption = "CNPJ:"
    Txt_Tarja.Text = "NÃO HÁ"
    'dados para o contador de emissoes
    var_SQL = "select * FROM tab_Trans_cartao "
    var_SQL = var_SQL + "WHERE cnpj_fornecedor = '" & txt_Codigo.Text & "'"
    var_SQL = var_SQL + "and veiculo = '" & CMB_VEICULO.Text & "'"
    Set DS = New Recordset
    DS.Open var_SQL, cnConexao
    Txt_TotalEmitidos.Text = DS.RecordCount
ElseIf vgl_TipoResponsavel = "OM" Then
    lbl_Codigo.Caption = "Cod. OM:"
    If vgl_Cor = "Azul" Then
       frm_Cartao_Emitir.Txt_Tarja.Text = "AZUL"
    Else
       frm_Cartao_Emitir.Txt_Tarja.Text = "LILÁS"
    End If
    'dados para o contador de emissoes
    var_SQL = "select * FROM tab_Trans_cartao "
    var_SQL = var_SQL + "WHERE cod_om = '" & txt_Codigo.Text & "'"
    var_SQL = var_SQL + "and veiculo = '" & CMB_VEICULO.Text & "'"
    Set DS = New Recordset
    DS.Open var_SQL, cnConexao
    Txt_TotalEmitidos.Text = DS.RecordCount
End If
Call Txt_Tarja_LostFocus
End Sub

Private Sub Extraviado_Click()
If Extraviado.Value = 1 Then
   TXT_Motivo.Enabled = True
ElseIf Extraviado.Value = 0 Then
   TXT_Motivo.Text = ""
   TXT_Motivo.Enabled = False
End If
End Sub
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Txt_DtEmissao.Text = Str(Date)
txt_Ano.Text = Right(Date, 4)
Var_Graduacao = ""
If (vgl_placa <> "___-____") Then
   If vgl_placa <> "" Then
      SQL = "SELECT TIPO FROM TAB_TRANS_VEICULO,TAB_TRANS_MODELO_VEIC"
      SQL = SQL + " WHERE COD = COD_MODELO"
      SQL = SQL + " AND PLACA='" & vgl_placa & "'"
      Set DS = New Recordset
      DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      If DS.RecordCount <> 0 Then
         CMB_VEICULO = DS(0)
      End If
   End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
vgl_TipoResponsavel = ""
End Sub
Private Sub txt_Ano_LostFocus()
Call Txt_Tarja_LostFocus
End Sub
Private Sub txt_Autorizado_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txt_tarja_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub LIMPAR()
txt_Ano = ""
CMB_VEICULO = ""
txt_Autorizado = ""
Txt_Tarja = ""
Txt_DtEmissao = ""
Txt_TotalEmitidos = ""
CMB_VEICULO.Text = ""
Txt_DtEmissao.Text = ""
txt_Autorizado = ""
Txt_Validade = ""
txt_Codigo = ""
TXT_Motivo = ""
Extraviado.Value = 0
txt_Descricao = ""
End Sub
Private Sub txt_NumCartao_LostFocus()
Call Txt_Tarja_LostFocus
End Sub
Private Sub Txt_Tarja_LostFocus()
If CMB_VEICULO.Text = "" Then Exit Sub
var_SQL = "select * FROM tab_Trans_cartao "
var_SQL = var_SQL + "WHERE nr_cartao = '" & txt_NumCartao & "' "
var_SQL = var_SQL + "AND Tarja = '" & Txt_Tarja.Text & "'"
var_SQL = var_SQL + "AND ano = '" & txt_Ano.Text & "'"
var_SQL = var_SQL + "AND veiculo = '" & CMB_VEICULO & "'"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 1 Then
   VarAcao = "Registro Velho"
   txt_Ano = DS!Ano
   CMB_VEICULO = DS!Veiculo
   txt_Autorizado = DS!autorizacao
   Txt_Tarja = DS!TARJA
   Txt_DtEmissao = DS!DT_EMISSAO
   CMB_VEICULO.Text = DS!Veiculo
   txt_Autorizado = DS!autorizacao
   Txt_Validade = DS!data_validade
   If DS!Extraviado = "Sim" Then
      Extraviado.Value = 1
   ElseIf DS!Extraviado = "Não" Then
      Extraviado.Value = 0
   End If
   TXT_Motivo = DS!Motivo
   TXT_Motivo.Enabled = False
   Extraviado.Enabled = False
   If DS!COD_OM <> "" Then
      txt_Codigo.Text = DS!COD_OM
      SQL = "select * FROM CONS_CARTAO"
      SQL = SQL + " where NR_CARTAO='" & txt_NumCartao & "'"
      Set DS = New Recordset
      DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      txt_Descricao = DS!NOME
   ElseIf DS!CPF_pessoa <> "" Then
      txt_Codigo.Text = DS!CPF_pessoa
      SQL = "select * FROM CONS_CARTAO"
      SQL = SQL + " where NR_CARTAO='" & txt_NumCartao & "'"
      Set DS = New Recordset
      DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      txt_Descricao = DS!NOME
   ElseIf DS!cnpj_fornecedor <> "" Then
      txt_Codigo.Text = DS!cnpj_fornecedor
      SQL = "select * FROM CONS_CARTAO"
      SQL = SQL + " where NR_CARTAO='" & txt_NumCartao & "'"
      Set DS = New Recordset
      DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      txt_Descricao = DS!NOME
   End If
   
   btn_Salvar.Enabled = False
   btn_Editar.Enabled = True
   btn_Excluir.Enabled = True
   btn_Cartao.Enabled = True
   
Else
   VarAcao = "Registro Novo"
   If txt_Codigo.Text = "" Then Exit Sub
   FRA_CARTAO.Enabled = True
   txt_NumCartao.Enabled = True
   txt_Ano.Enabled = True
   CMB_VEICULO.Enabled = True
   txt_Autorizado.Enabled = True
   
   If Txt_Validade <> "" Then
      Txt_Validade.Enabled = True
   End If
   btn_Editar.Enabled = False
   btn_Excluir.Enabled = False
   btn_Cartao.Enabled = False
   btn_Salvar.Enabled = True
End If
End Sub
