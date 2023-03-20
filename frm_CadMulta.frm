VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_CadMulta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SisTrans - Cadastro de Multas"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8745
   ControlBox      =   0   'False
   Icon            =   "frm_CadMulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   135
      Left            =   120
      TabIndex        =   27
      Top             =   4200
      Width           =   8535
   End
   Begin VB.CommandButton btn_Salvar 
      Caption         =   "Sal&var"
      Height          =   855
      Left            =   120
      Picture         =   "frm_CadMulta.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Salvar"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton btn_Editar 
      Caption         =   "E&ditar"
      Height          =   855
      Left            =   1080
      Picture         =   "frm_CadMulta.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Editar"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton btn_Excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   2040
      Picture         =   "frm_CadMulta.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Excluir"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton btn_Sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   7800
      Picture         =   "frm_CadMulta.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   4440
      Width           =   855
   End
   Begin VB.Frame fra_Multa 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   240
      TabIndex        =   23
      Top             =   360
      Width           =   8055
      Begin MSMask.MaskEdBox txt_NumTiquete 
         Height          =   300
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   4
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt_Data 
         Height          =   300
         Left            =   2760
         TabIndex        =   1
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Frame fra_Veiculo 
         Caption         =   "Veículo:"
         Enabled         =   0   'False
         Height          =   2055
         Left            =   0
         TabIndex        =   31
         Top             =   1560
         Width           =   8055
         Begin VB.Frame Frame1 
            Caption         =   "Dados:"
            Height          =   855
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   7815
            Begin MSMask.MaskEdBox txt_Placa 
               Height          =   300
               Left            =   120
               TabIndex        =   6
               Top             =   480
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   529
               _Version        =   393216
               BackColor       =   16777215
               MaxLength       =   8
               Mask            =   "AAA-9999"
               PromptChar      =   "_"
            End
            Begin VB.TextBox txt_Modelo 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   480
               Width           =   3855
            End
            Begin VB.TextBox txt_Cor 
               Enabled         =   0   'False
               Height          =   285
               Left            =   5520
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   480
               Width           =   2175
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Placa:"
               Height          =   195
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   450
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Modelo:"
               Height          =   195
               Left            =   1560
               TabIndex        =   34
               Top             =   240
               Width           =   570
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Cor:"
               Height          =   195
               Left            =   5520
               TabIndex        =   33
               Top             =   240
               Width           =   285
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Responsável:"
            Height          =   855
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Width           =   7815
            Begin VB.TextBox txt_Tipo 
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               MaxLength       =   15
               TabIndex        =   9
               Text            =   "CPF/CNPJ/OM"
               Top             =   480
               Width           =   2055
            End
            Begin VB.TextBox txt_NomeResp 
               Enabled         =   0   'False
               Height          =   285
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   10
               Top             =   480
               Width           =   5415
            End
            Begin VB.Label lbl_Tipo 
               AutoSize        =   -1  'True
               Caption         =   "Tipo:"
               Height          =   195
               Left            =   120
               TabIndex        =   38
               Top             =   240
               Width           =   360
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Descrição:"
               Height          =   195
               Left            =   2280
               TabIndex        =   37
               Top             =   240
               Width           =   765
            End
         End
      End
      Begin VB.Frame fra_DadosMulta 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1575
         Left            =   2760
         TabIndex        =   39
         Top             =   120
         Width           =   5055
         Begin VB.Frame Frame4 
            Height          =   135
            Left            =   0
            TabIndex        =   46
            Top             =   600
            Width           =   4935
         End
         Begin VB.TextBox txt_Hora 
            Height          =   285
            Left            =   1320
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txt_Local 
            Height          =   285
            Left            =   2040
            TabIndex        =   3
            Top             =   240
            Width           =   2895
         End
         Begin MSDataListLib.DataCombo dbc_TipoMulta 
            Height          =   315
            Left            =   0
            TabIndex        =   4
            Top             =   1080
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Desc_Tipo_Multa"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_Controlador 
            Height          =   315
            Left            =   3240
            TabIndex        =   5
            Top             =   1080
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            BackColor       =   -2147483648
            ListField       =   "codinome"
            Text            =   ""
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Controlador de Trânsito:"
            Height          =   195
            Left            =   3240
            TabIndex        =   45
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Multa:"
            Height          =   195
            Left            =   0
            TabIndex        =   43
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   390
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   1320
            TabIndex        =   41
            Top             =   0
            Width           =   390
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Local:"
            Height          =   195
            Left            =   2040
            TabIndex        =   40
            Top             =   0
            Width           =   435
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nº do Tíquete de infração:"
         Height          =   195
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   1905
      End
      Begin VB.Image imgArm 
         Height          =   645
         Index           =   0
         Left            =   360
         Picture         =   "frm_CadMulta.frx":1412
         Stretch         =   -1  'True
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Frame fra_ListaMulta 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   240
      TabIndex        =   44
      Top             =   480
      Width           =   8175
      Begin MSDataGridLib.DataGrid dbg_Listagem 
         Height          =   3375
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5953
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
      Height          =   3255
      Left            =   240
      TabIndex        =   25
      Top             =   360
      Width           =   8055
      Begin VB.TextBox Txt_Requerimento 
         Height          =   1815
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   960
         Width           =   7575
      End
      Begin VB.TextBox txt_NomeRequer 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txt_OrgExp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txt_RG 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chk_Proprio 
         Caption         =   "O próprio"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame Frame6 
         Caption         =   "CGS:"
         Enabled         =   0   'False
         Height          =   855
         Left            =   6720
         TabIndex        =   26
         Top             =   0
         Width           =   1335
         Begin VB.OptionButton opt_Indeferido 
            Caption         =   "Indeferido"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton opt_Deferido 
            Caption         =   "Deferido"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nome do requerente:"
         Height          =   195
         Left            =   1320
         TabIndex        =   30
         Top             =   120
         Width           =   1500
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Orgão Exp.:"
         Height          =   195
         Left            =   5640
         TabIndex        =   29
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "RG:"
         Height          =   195
         Left            =   3960
         TabIndex        =   28
         Top             =   120
         Width           =   285
      End
   End
   Begin MSComctlLib.TabStrip TabSctripMenu 
      Height          =   4215
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7435
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
            Object.ToolTipText     =   "Histórico de Multas"
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
Attribute VB_Name = "frm_CadMulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Editar_Click()

fra_DadosMulta.Enabled = True
fra_Veiculo.Enabled = True
fra_Requerimento.Enabled = True

btn_Salvar.Enabled = True
btn_Editar.Enabled = False
btn_Excluir.Enabled = False

End Sub
Private Sub btn_Excluir_Click()

On Error Resume Next

If MsgBox("Deseja excluir o registro?", vbYesNo + vbQuestion, "SisTrans") = vbYes Then
    
    If txt_NumTiquete.Text = "" Then
        MsgBox "Não há registro para excluir,", vbOKOnly + vbInformation, "SisTrans"
        txt_Placa.SetFocus
        Exit Sub
    End If
    rs_tabMulta.Close
    
    Set rs_tabMulta = New Recordset
    rs_tabMulta.Open "delete * from tab_Trans_multa where nr_talao_infr ='" & txt_NumTiquete.Text & "'", cnConexao, adOpenStatic, adLockOptimistic
    
    MsgBox "Arquivo excluído.", vbOKOnly + vbInformation, "SisTrans"
    
    Call LIMPAR
    txt_NumTiquete = ""
    Call Form_Load
    
End If
Exit Sub
Error:
    MsgBox "Erro. O registro não foi excluído.", vbOKOnly + vbInformation, "SisTrans"
    btn_Sair.SetFocus
End Sub
Private Sub btn_Sair_Click()
Unload Me
End Sub
Private Sub btn_Salvar_Click()

If dbc_TipoMulta.Text = "" Then
   MsgBox "Selecione o tipo da multa.", vbOKOnly + vbInformation, "SisTrans"
   dbc_TipoMulta.SetFocus
    Exit Sub
End If
If txt_Placa = "___-____" Then
   MsgBox "Digite a placa.", vbOKOnly + vbInformation, "SisTrans"
   txt_Placa.SetFocus
   Exit Sub
End If

SQL = "select * from tab_trans_Multa"
SQL = SQL + " where nr_talao_infr ='" & txt_NumTiquete.Text & "'"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
With DS
     If .RecordCount = 0 Then
        .AddNew
        !nr_talao_infr = txt_NumTiquete.Text
     End If
    
     !placa = txt_Placa.Text
    
     !dt_infr = txt_Data.Text
     !hr_infr = txt_Hora.Text
     !Local = txt_Local.Text
     !Dt_Listagem_Detran = ""
     !deferido = "-"
     !nome_requerente = txt_NomeRequer.Text
     !rg_requerente = txt_RG.Text
     !orgao_exp = txt_OrgExp.Text
     !Detran = "SIM"
            
     !req_CGS = Txt_Requerimento
    
     '---------------------------
     'para obter o cod_tipo_multa
     '---------------------------
     SQL = "select cod from TAB_TRANS_TIPO_MULTA"
     SQL = SQL + " where desc_tipo_multa='" & dbc_TipoMulta.Text & "'"
     Set ds1 = New Recordset
     ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
    
     VarCodMulta = ds1(0)
     
     !cod_tipo_multa = VarCodMulta
     
     .UpdateBatch adAffectAll
     MsgBox "Arquivo salvo.", vbOKOnly + vbInformation, "SisTrans"
        
     TabSctripMenu.Tabs(1).Selected = True
End With

Call LIMPAR
Call Form_Load
txt_NumTiquete.Text = ""

End Sub

Private Sub Form_Load()
'----------
Me.Top = 0
Me.Left = 0
'----------
Call LIMPAR
'-----------------------------------
'Enchendo a combo de tipos de multa.
'-----------------------------------
SQL = "select * from tab_Trans_tipo_multa"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set dbc_TipoMulta.RowSource = DS
'-----------------------------------

End Sub
Private Sub TabSctripMenu_Click()
Select Case TabSctripMenu.SelectedItem.Index
    Case 1
          fra_Multa.Visible = True
          fra_Requerimento.Visible = False
          fra_ListaMulta.Visible = False
    Case 2
          fra_Multa.Visible = False
          fra_Requerimento.Visible = False
          fra_ListaMulta.Visible = True
    Case 3
          fra_Multa.Visible = False
          fra_Requerimento.Visible = True
          fra_ListaMulta.Visible = False
End Select
End Sub
Private Sub txt_Local_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_NumTiquete_Change()
Call LIMPAR
End Sub

Private Sub txt_NumTiquete_LostFocus()
If txt_NumTiquete.Text = "" Then Exit Sub
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
              btn_Editar.Enabled = False
           ElseIf !deferido = "DEFERIDO" Then
              opt_Deferido.Value = True
              btn_Editar.Enabled = False
           ElseIf !deferido = "-" Then
              opt_Deferido.Value = False
              opt_Indeferido.Value = False
              btn_Editar.Enabled = True
           End If
           '--------------------------
           txt_NomeResp.Text = "Responsável não identificado"
           If txt_NomeResp = !nome_requerente And txt_NomeResp <> "O responsavel não foi identificado" Then
              chk_Proprio.Value = 1
           End If
           chk_Proprio.Enabled = False
           txt_NomeRequer.Enabled = False
           txt_RG.Enabled = False
           txt_OrgExp.Enabled = False
           txt_Data.Enabled = False
           
           txt_NomeRequer = !nome_requerente
           txt_RG = !rg_requerente
           txt_OrgExp = !orgao_exp
      
           fra_DadosMulta.Enabled = False
           fra_Veiculo.Enabled = False
           fra_Requerimento.Enabled = False
                
           btn_Salvar.Enabled = False
        
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
      txt_Data.Text = Str(Date)
      txt_Hora.Text = Str(Time)
      fra_DadosMulta.Enabled = True
      fra_Veiculo.Enabled = True
      fra_Requerimento.Enabled = True
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
           btn_Editar.Enabled = False
        ElseIf !deferido = "DEFERIDO" Then
            opt_Deferido.Value = True
            btn_Editar.Enabled = False
        ElseIf !deferido = "-" Then
            opt_Deferido.Value = False
            opt_Indeferido.Value = False
            btn_Editar.Enabled = True
        End If
        '--------------------------
               
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
         If txt_NomeResp = !nome_requerente And txt_NomeResp <> "O responsavel não foi identificado" Then
            chk_Proprio.Value = 1
         End If
         chk_Proprio.Enabled = False
         txt_NomeRequer.Enabled = False
         txt_RG.Enabled = False
         txt_OrgExp.Enabled = False
         txt_Data.Enabled = False
         
         txt_NomeRequer = !nome_requerente
         txt_RG = !rg_requerente
         txt_OrgExp = !orgao_exp
      
         fra_DadosMulta.Enabled = False
         fra_Veiculo.Enabled = False
         fra_Requerimento.Enabled = False
                
         btn_Salvar.Enabled = False
        
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
dbc_Controlador.Text = ""
txt_NomeRequer.Text = ""
txt_RG.Text = ""
txt_OrgExp.Text = ""
txt_Data = ""
txt_Hora.Text = ""
txt_Local.Text = ""

chk_Proprio.Value = 0
opt_Indeferido.Value = False
opt_Deferido.Value = False
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
If DS.RecordCount = 0 Then
   SQL = "select * from cns_Trans_veiculo_null,tab_trans_modelo_veic,"
   SQL = SQL + " tab_trans_marca_veic where "
   SQL = SQL + " cns_trans_veiculo_null.placa = '" & txt_Placa.Text & "'"
   SQL = SQL + " and cns_trans_veiculo_null.cod_modelo=tab_trans_modelo_veic.cod"
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
            
              SQL = "select * from cns_Trans_multa_todos_dados"
              SQL = SQL + " where "
              SQL = SQL + " nr_talao_infr = '" & txt_NumTiquete.Text & "'"
              Set DS = New Recordset
              DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
              If DS.RecordCount <> 0 Then
                 txt_RG.Text = DS!rg_requerente
                 txt_OrgExp.Text = DS!orgao_exp
                 txt_NomeRequer = DS!nome_requerente
              End If
           ElseIf !COD_OM <> "" Then
              lbl_Tipo.Caption = "Sigla:"
              TXT_TIPO.Text = !COD_OM
              txt_NomeResp.Text = !NOME
           Else
              txt_NomeResp.Text = "Responsável não identificado"
           End If
        Else
           frm_CadVeiculo.txt_Placa = txt_Placa
        End If
    End With
Else
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
            
               SQL = "select * from cns_Trans_multa_todos_dados"
               SQL = SQL + " where "
               SQL = SQL + " nr_talao_infr = '" & txt_NumTiquete.Text & "'"
               Set DS = New Recordset
               DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
               If DS.RecordCount <> 0 Then
                  txt_RG.Text = DS!rg_requerente
                  txt_OrgExp.Text = DS!orgao_exp
                  txt_NomeRequer = DS!nome_requerente
               End If
            ElseIf !COD_OM <> "" Then
               lbl_Tipo.Caption = "Sigla:"
               TXT_TIPO.Text = !COD_OM
               txt_NomeResp.Text = !NOME
            End If
         Else
           frm_CadVeiculo.txt_Placa = txt_Placa
         End If
    End With
End If
End Sub

Private Sub Txt_Requerimento_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
