VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_CadMultaReq 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SisTrans - Requerimento de Multas"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8535
   ControlBox      =   0   'False
   Icon            =   "frm_CadMultaReq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_CGS 
      Caption         =   "CGS:"
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
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   6000
      TabIndex        =   43
      Top             =   3240
      Width           =   2295
      Begin VB.OptionButton opt_Deferido 
         Caption         =   "Deferido"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton opt_Indeferido 
         Caption         =   "Indeferido"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Dado editado apenas pelo Sr CGS."
         Height          =   435
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   1965
      End
   End
   Begin VB.Frame Frame3 
      Height          =   135
      Left            =   120
      TabIndex        =   24
      Top             =   4800
      Width           =   8295
   End
   Begin VB.CommandButton btn_Salvar 
      Caption         =   "Sal&var"
      Height          =   855
      Left            =   120
      Picture         =   "frm_CadMultaReq.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Salvar"
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton btn_Sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   7560
      Picture         =   "frm_CadMultaReq.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   5040
      Width           =   855
   End
   Begin VB.Frame fra_Multa 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   240
      TabIndex        =   20
      Top             =   360
      Width           =   8055
      Begin MSMask.MaskEdBox txt_NumTiquete 
         Height          =   300
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         PromptChar      =   "_"
      End
      Begin VB.Frame fra_DadosReq 
         Caption         =   "Dados do Requerimento:"
         Height          =   1335
         Left            =   0
         TabIndex        =   37
         Top             =   2880
         Width           =   5655
         Begin VB.CheckBox chk_Proprio 
            Caption         =   "O próprio (Caso o Requerente seja o responsável do veículo):"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   4695
         End
         Begin VB.TextBox txt_RG 
            Height          =   285
            Left            =   2760
            TabIndex        =   12
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txt_OrgExp 
            Height          =   285
            Left            =   4440
            TabIndex        =   13
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txt_NomeRequer 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "RG:"
            Height          =   195
            Left            =   2760
            TabIndex        =   40
            Top             =   600
            Width           =   285
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Orgão Exp.:"
            Height          =   195
            Left            =   4440
            TabIndex        =   39
            Top             =   600
            Width           =   840
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Nome do requerente:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   600
            Width           =   1500
         End
      End
      Begin VB.Frame fra_DadosMulta 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   615
         Left            =   2280
         TabIndex        =   1
         Top             =   120
         Width           =   5655
         Begin VB.TextBox txt_Hora 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txt_Local 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
         Begin MSDataListLib.DataCombo dbc_TipoMulta 
            Height          =   315
            Left            =   3600
            TabIndex        =   4
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "Desc_Tipo_Multa"
            Text            =   ""
         End
         Begin MSMask.MaskEdBox txt_Data 
            Height          =   300
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Multa:"
            Height          =   195
            Left            =   3840
            TabIndex        =   36
            Top             =   0
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   0
            Width           =   390
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   1320
            TabIndex        =   34
            Top             =   0
            Width           =   390
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Local:"
            Height          =   195
            Left            =   2040
            TabIndex        =   33
            Top             =   0
            Width           =   435
         End
      End
      Begin VB.Frame fra_Veiculo 
         Caption         =   "Veículo:"
         Enabled         =   0   'False
         Height          =   2175
         Left            =   0
         TabIndex        =   25
         Top             =   720
         Width           =   8055
         Begin VB.Frame Frame2 
            Caption         =   "Responsável:"
            Height          =   855
            Left            =   120
            TabIndex        =   30
            Top             =   1200
            Width           =   7815
            Begin VB.TextBox txt_Tipo 
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   8
               Text            =   "CPF/CNPJ/OM"
               Top             =   480
               Width           =   2055
            End
            Begin VB.TextBox txt_NomeResp 
               Enabled         =   0   'False
               Height          =   285
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   480
               Width           =   5415
            End
            Begin VB.Label lbl_Tipo 
               AutoSize        =   -1  'True
               Caption         =   "Tipo:"
               Height          =   195
               Left            =   120
               TabIndex        =   32
               Top             =   240
               Width           =   360
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Descrição:"
               Height          =   195
               Left            =   2280
               TabIndex        =   31
               Top             =   240
               Width           =   765
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Dados:"
            Height          =   855
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   5775
            Begin MSMask.MaskEdBox txt_Placa 
               Height          =   300
               Left            =   120
               TabIndex        =   5
               Top             =   480
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   8
               Mask            =   "AAA-9999"
               PromptChar      =   "_"
            End
            Begin VB.TextBox txt_Modelo 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   480
               Width           =   2535
            End
            Begin VB.TextBox txt_Cor 
               Enabled         =   0   'False
               Height          =   285
               Left            =   4200
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Placa:"
               Height          =   195
               Left            =   120
               TabIndex        =   29
               Top             =   240
               Width           =   450
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Modelo:"
               Height          =   195
               Left            =   1560
               TabIndex        =   28
               Top             =   120
               Width           =   570
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Cor:"
               Height          =   195
               Left            =   4200
               TabIndex        =   27
               Top             =   240
               Width           =   285
            End
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nº do Tíquete de infração:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   1905
      End
   End
   Begin VB.Frame fra_ListaMulta 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   240
      TabIndex        =   41
      Top             =   360
      Width           =   7695
      Begin MSDataGridLib.DataGrid dbg_Listagem 
         Height          =   4095
         Left            =   360
         TabIndex        =   18
         Top             =   120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7223
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
         Caption         =   "Histórico de multas do veículo:"
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
   Begin VB.Frame fra_Requerimento 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   240
      TabIndex        =   23
      Top             =   360
      Width           =   8055
      Begin VB.Frame Fra_REQ 
         Caption         =   "Requerimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   7695
         Begin VB.TextBox Txt_Requerimento 
            Height          =   1815
            Left            =   360
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   480
            Width           =   7095
         End
      End
   End
   Begin MSComctlLib.TabStrip TabSctripMenu 
      Height          =   4815
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8493
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Informações:"
            Object.ToolTipText     =   "Entre com os dados da multa"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Multas:"
            Object.ToolTipText     =   "Lista de Multas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Requerimento:"
            Object.ToolTipText     =   "Insira o requerimento escaniado"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frm_CadMultaReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Sair_Click()
Unload Me
End Sub
Private Sub btn_Salvar_Click()
If txt_NumTiquete = "_________" Then Exit Sub
'Abre a tabela verifica se a senha esta correta, caso sim edita-a

SQL = "select * from tab_trans_Multa"
SQL = SQL + " where nr_talao_infr ='" & txt_NumTiquete.Text & "'"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
With DS
     If Txt_Requerimento <> "" Then
        If txt_NomeRequer = "" Then
           MsgBox "Cadastre o requerente", vbOKOnly + vbInformation, "SisTrans"
           Exit Sub
        Else
          !req_CGS = Txt_Requerimento
          !nome_requerente = txt_NomeRequer.Text
          !rg_requerente = txt_RG.Text
          !orgao_exp = txt_OrgExp.Text
    
          If vgl_Nivel = "CGS" Then
             If opt_Indeferido.Value = True Then
                !deferido = "INDEFERIDO"
                !Detran = "SIM"
             ElseIf opt_Deferido.Value = True Then
                !deferido = "DEFERIDO"
                !Detran = "NÃO"
             End If
          End If
          .UpdateBatch adAffectAll
          MsgBox "Arquivo salvo.", vbOKOnly + vbInformation, "SisTrans"
        End If
     Else
        MsgBox "Digite o requerimento.", vbOKOnly + vbInformation, "SisTrans"
     End If
End With
Call LIMPAR
txt_NumTiquete.Text = ""
Call Form_Load
End Sub
Private Sub chk_Proprio_Click()
If chk_Proprio.Value = 1 Then
   
   SQL = "select * from cns_Trans_veiculo,tab_trans_modelo_veic,"
   SQL = SQL + " tab_trans_marca_veic,tab_ger_pessoa where "
   SQL = SQL + " cns_trans_veiculo.placa = '" & txt_Placa.Text & "'"
   SQL = SQL + " and cns_trans_veiculo.cod_modelo=tab_trans_modelo_veic.cod"
   SQL = SQL + " and tab_trans_modelo_veic.cod_marca=tab_trans_marca_veic.cod"
   SQL = SQL + " and tab_trans_veiculo.cpf_resp_pessoa=tab_ger_pessoa.cpf"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   If DS.RecordCount <> 0 Then
      txt_RG.Text = DS!identidade
      txt_OrgExp.Text = DS!OrgaoExped_Id
      txt_NomeRequer = DS!NOME
   End If
   
   txt_NomeRequer.Enabled = False
   txt_RG.Enabled = False
   txt_OrgExp.Enabled = False
Else
   txt_NomeRequer.Enabled = True
   txt_RG.Enabled = True
   txt_OrgExp.Enabled = True
   txt_NomeRequer = ""
   txt_RG = ""
   txt_OrgExp = ""
End If
End Sub

Private Sub dbg_Listagem_Click()
SQL = "select * from cns_Trans_Multa_todos_dados"
SQL = SQL + " where nr_talao_infr ='" & dbg_Listagem.Columns(0) & "'"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

fra_Multa.Visible = False
fra_ListaMulta.Visible = True

txt_NumTiquete = dbg_Listagem.Columns(0)
Call txt_NumTiquete_LostFocus
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Call LIMPAR
SQL = "select * from tab_Trans_tipo_multa"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set dbc_TipoMulta.RowSource = DS

If vgl_Nivel = "CGS" Then
   TabSctripMenu.Tabs(3).Selected = True
   fra_DadosMulta.Enabled = False
   fra_Veiculo.Enabled = False
   Fra_REQ.Enabled = False
   fra_CGS.Enabled = True
   Call txt_NumTiquete_LostFocus
Else
   fra_DadosMulta.Enabled = False
   fra_Veiculo.Enabled = False
   Fra_REQ.Enabled = True
   fra_CGS.Enabled = False
End If
End Sub
Private Sub TabSctripMenu_Click()
Select Case TabSctripMenu.SelectedItem.Index
       Case 1
            fra_CGS.Visible = True
            fra_Multa.Visible = True
            fra_Requerimento.Visible = False
            fra_ListaMulta.Visible = False
       Case 2
            fra_Multa.Visible = False
            fra_Requerimento.Visible = False
            fra_ListaMulta.Visible = True
            fra_CGS.Visible = False
       Case 3
            fra_Multa.Visible = False
            fra_Requerimento.Visible = True
            fra_ListaMulta.Visible = False
            fra_CGS.Visible = False
End Select
End Sub
Private Sub txt_NomeRequer_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Public Sub txt_NumTiquete_LostFocus()
If txt_NumTiquete.Text = "_________" Then Exit Sub

SQL = "select * from cns_Trans_Multa_todos_dados,cns_trans_multa"
SQL = SQL + " where cns_Trans_Multa_todos_dados.nr_talao_infr ='" & txt_NumTiquete.Text & "'"
SQL = SQL + " and cns_Trans_Multa_todos_dados.nr_talao_infr=cns_trans_multa.nr_talao_infr"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
   SQL = "select * from cns_Trans_Multa_todos_dados"
   SQL = SQL + " where cns_Trans_Multa_todos_dados.nr_talao_infr ='" & txt_NumTiquete.Text & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   If DS.RecordCount <> 0 Then
      With DS
           txt_Placa.Text = !placa
           txt_Data.Text = !dt_infr
           txt_Hora.Text = !hr_infr
           txt_Local.Text = !Local
           Txt_Requerimento = !req_CGS
           dbc_TipoMulta.Text = !desc_tipo_multa
           TXT_MODELO.Text = !MARCA & " - " & !modelo
           txt_Cor = !Cor_Pred
           
           '--------------------------
           'Trata se indeferido ou não
           '--------------------------
           If !deferido = "INDEFERIDO" Then
               opt_Indeferido.Value = True
           ElseIf !deferido = "DEFERIDO" Then
               opt_Deferido.Value = True
           ElseIf !deferido = "-" Then
               opt_Deferido.Value = False
               opt_Indeferido.Value = False
           End If
           '--------------------------
           
           txt_NomeResp.Text = "Responsável não identificado"
           txt_NomeRequer.Enabled = True
           txt_RG.Enabled = True
           txt_OrgExp.Enabled = True
   
           txt_NomeRequer = !nome_requerente
           txt_RG = !rg_requerente
           txt_OrgExp = !orgao_exp
         
           If vgl_Nivel = "CGS" Then
              TabSctripMenu.Tabs(3).Selected = True
              If !deferido <> "-" Then
                 Txt_Requerimento.Enabled = False
                 fra_DadosReq.Enabled = False
              Else
                 Txt_Requerimento.Enabled = True
                 fra_DadosReq.Enabled = True
              End If
              fra_DadosMulta.Enabled = False
              fra_Veiculo.Enabled = False
              Fra_REQ.Enabled = False
              fra_CGS.Enabled = True
           Else
              fra_DadosReq.Enabled = True
              fra_DadosMulta.Enabled = False
              fra_Veiculo.Enabled = False
              Fra_REQ.Enabled = True
              fra_CGS.Enabled = False
           End If
           '---------------------------------------
           'lança o historico de multas em listagem
           '---------------------------------------
           SQL = "select * from cns_Trans_Multa_todos_dados"
           SQL = SQL + " where tab_trans_multa.placa ='" & DS!placa & "'"
           Set DS = New Recordset
           DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
        
           Set dbg_Listagem.DataSource = DS
           '---------------------------------------
         End With
       Else
           Call LIMPAR
           txt_Data.Text = Date
           txt_Hora.Text = Time
        
           fra_DadosMulta.Enabled = True
           fra_Veiculo.Enabled = True
           Fra_REQ.Enabled = True
           fra_CGS.Enabled = False
           btn_Salvar.Enabled = True
       End If
Else
   With DS
        txt_Placa.Text = !placa
        txt_Data.Text = !dt_infr
        txt_Hora.Text = !hr_infr
        txt_Local.Text = !Local
        Txt_Requerimento = !req_CGS
        dbc_TipoMulta.Text = !desc_tipo_multa
        TXT_MODELO.Text = !MARCA & " - " & !modelo
        txt_Cor = !Cor_Pred
        '--------------------------
        'Trata se indeferido ou não
        '--------------------------
        If !deferido = "INDEFERIDO" Then
           opt_Indeferido.Value = True
        ElseIf !deferido = "DEFERIDO" Then
            opt_Deferido.Value = True
        ElseIf !deferido = "-" Then
            opt_Deferido.Value = False
            opt_Indeferido.Value = False
        End If
        '-----------------------------------
        'indica que o requerente é o próprio
        '-----------------------------------
             
        If !cnpj <> "" Then
           lbl_Tipo.Caption = "CNPJ:"
           TXT_TIPO.Text = !cnpj
           txt_NomeResp.Text = !NOME
        ElseIf !cpf_resp_pessoa <> "" Then
           lbl_Tipo.Caption = "CPF:"
           TXT_TIPO.Text = !cpf_resp_pessoa
           txt_NomeResp.Text = !NOME
        ElseIf !COD_OM <> "" Then
           lbl_Tipo.Caption = "Sigla:"
           TXT_TIPO.Text = !COD_OM
           txt_NomeResp.Text = !NOME
        End If
        If txt_NomeResp = !nome_requerente Then
           chk_Proprio.Value = 1
           txt_NomeRequer.Enabled = False
           txt_RG.Enabled = False
           txt_OrgExp.Enabled = False
        End If
        txt_NomeRequer.Enabled = True
        txt_RG.Enabled = True
        txt_OrgExp.Enabled = True
   
        txt_NomeRequer = !nome_requerente
        txt_RG = !rg_requerente
        txt_OrgExp = !orgao_exp
         
        If vgl_Nivel = "CGS" Then
           TabSctripMenu.Tabs(3).Selected = True
           If !deferido <> "-" Then
              Txt_Requerimento.Enabled = False
              fra_DadosReq.Enabled = False
           Else
             Txt_Requerimento.Enabled = True
             fra_DadosReq.Enabled = True
           End If
           fra_DadosMulta.Enabled = False
           fra_Veiculo.Enabled = False
           Fra_REQ.Enabled = False
           fra_CGS.Enabled = True
        Else
           If !deferido <> "-" Then
              Txt_Requerimento.Enabled = False
              fra_DadosReq.Enabled = False
           Else
              Txt_Requerimento.Enabled = True
              fra_DadosReq.Enabled = True
           End If
           fra_DadosMulta.Enabled = False
           fra_Veiculo.Enabled = False
           Fra_REQ.Enabled = True
           fra_CGS.Enabled = False
        End If
        '---------------------------------------
        'lança o historico de multas em listagem
        '---------------------------------------
        SQL = "select * from cns_Trans_Multa_todos_dados"
        SQL = SQL + " where tab_trans_multa.placa ='" & DS!placa & "'"
        Set DS = New Recordset
        DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
        
        Set dbg_Listagem.DataSource = DS
        '---------------------------------------
    End With
End If
End Sub
Private Sub LIMPAR()
txt_Placa.Mask = "        "
txt_Placa.Mask = "AAA-9999"
TXT_MODELO.Text = ""
txt_Cor.Text = ""

lbl_Tipo.Caption = "Código:"
TXT_TIPO.Text = ""
txt_NomeResp.Text = ""
dbc_TipoMulta.Text = ""

txt_NomeRequer.Text = ""
txt_RG.Text = ""
txt_OrgExp.Text = ""
txt_Data.Mask = "          "
txt_Data.Mask = "99/99/9999"
txt_Hora.Text = ""
txt_Local.Text = ""
Txt_Requerimento = ""

chk_Proprio.Value = 0

opt_Indeferido.Value = False
opt_Deferido.Value = False

End Sub
Private Sub txt_OrgExp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txt_Placa_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txt_Placa_LostFocus()

SQL = "select * from cns_Trans_veiculo,tab_trans_modelo_veic,"
SQL = SQL + " tab_trans_marca_veic where "
SQL = SQL + " cns_trans_veiculo.placa = '" & txt_Placa.Text & "'"
SQL = SQL + " and cns_trans_veiculo.cod_modelo=tab_trans_modelo_veic.cod"
SQL = SQL + " and tab_trans_modelo_veic.cod_marca=tab_trans_marca_veic.cod"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
With DS
     If .RecordCount <> 0 Then
     
        TXT_MODELO.Text = !MARCA & " - " & !modelo
        txt_Cor = !Cor_Pred
        
        If !cnpj <> "" Then
            lbl_Tipo.Caption = "CNPJ:"
            TXT_TIPO.Text = !cnpj
            txt_NomeResp.Text = !NOME
        ElseIf !cpf_resp_pessoa <> "" Then
            lbl_Tipo.Caption = "CPF:"
            TXT_TIPO.Text = !cpf_resp_pessoa
            txt_NomeResp.Text = !NOME
            
       ElseIf !COD_OM <> "" Then
                           
            lbl_Tipo.Caption = "Sigla:"
            TXT_TIPO.Text = !COD_OM
            txt_NomeResp.Text = !NOME
        End If
    Else
        frm_CadVeiculo.txt_Placa = txt_Placa
    End If
End With
End Sub
Private Sub Txt_Requerimento_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
