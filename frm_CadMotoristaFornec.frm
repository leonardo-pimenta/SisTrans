VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_CadMotoristaFornec 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SisTrans - Cadastro de Motoristas - Fornecedor"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6690
   ControlBox      =   0   'False
   Icon            =   "frm_CadMotoristaFornec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_CNPJ 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Height          =   135
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   6495
   End
   Begin VB.CommandButton btn_Salvar 
      Caption         =   "Sal&var"
      Height          =   855
      Left            =   120
      Picture         =   "frm_CadMotoristaFornec.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salvar"
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton btn_Excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1080
      Picture         =   "frm_CadMotoristaFornec.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Excluir"
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton btn_Sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5760
      Picture         =   "frm_CadMotoristaFornec.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   2400
      Width           =   855
   End
   Begin VB.Frame fra_Cadastro 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txt_NomeFornec 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   4695
      End
      Begin VB.TextBox txt_NomeMotorista 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   0
         TabIndex        =   5
         Top             =   1680
         Width           =   6495
      End
      Begin VB.Frame Fra_Identidade 
         Caption         =   "Identidade:"
         Height          =   615
         Left            =   0
         TabIndex        =   11
         Top             =   720
         Width           =   6495
         Begin VB.TextBox txt_OrgaoExp 
            Height          =   285
            Left            =   3360
            TabIndex        =   3
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txt_Identidade 
            Height          =   285
            Left            =   360
            MaxLength       =   15
            TabIndex        =   2
            Top             =   240
            Width           =   1935
         End
         Begin MSMask.MaskEdBox txt_Validade 
            Height          =   300
            Left            =   5160
            TabIndex        =   4
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nº:"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   225
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Orgão Exp.:"
            Height          =   195
            Left            =   2400
            TabIndex        =   13
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Validade:"
            Height          =   195
            Left            =   4440
            TabIndex        =   12
            Top             =   240
            Width           =   660
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Fornecedor:"
         Height          =   195
         Left            =   1800
         TabIndex        =   17
         Top             =   120
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ:"
         Height          =   195
         Left            =   0
         TabIndex        =   16
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Motorista:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frm_CadMotoristaFornec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public varModoAbrir As Byte
Public varCNPJ As String
Public varCodMotorista As String
Dim rs_tabMotoristaFornec As Recordset
Private Sub btn_Excluir_Click()

On Error GoTo Error

If MsgBox("Deseja excluir o registro?", vbYesNo + vbQuestion, "SisTrans") = vbYes Then
    
    If TXT_IDENTIDADE = "" Then
        MsgBox "Não há registro para excluir,", vbOKOnly + vbInformation, "SisTrans"
        Exit Sub
    End If
    var_SQL = "DELETE * FROM tab_Trans_Motorista_Fornecedor "
    var_SQL = var_SQL + " where identidade ='" & TXT_IDENTIDADE.Text & "'"
    Set DS = New Recordset
    DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
    
    MsgBox "Arquivo excluído.", vbOKOnly + vbInformation, "SisTrans"
    
    Call prc_LimparCampos
    
    Unload Me
    
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

On Error GoTo Erro

If txt_NomeMotorista.Text = "" Then
    MsgBox "Entre com o nome do motorista", vbInformation + vbOKOnly, "SisTrans"
    Exit Sub
End If
'Abre a tabela verifica se a senha esta correta, caso sim edita-a
Set rs_tabMotoristaFornec = New Recordset
With rs_tabMotoristaFornec

    If varCodMotorista = "" Then varCodMotorista = 0
    
    .Open "select * from tab_trans_motorista_fornecedor where cod_motorista =" & varCodMotorista & "", cnConexao, adOpenStatic, adLockOptimistic

    If .RecordCount = 0 Then
        Set rs_tab = New Recordset
        With rs_tab
            .Open "select * from tab_trans_motorista_fornecedor where cnpj_FORNEC = '" & txt_CNPJ.Text & "'", cnConexao, adOpenStatic, adLockOptimistic
            
            If .RecordCount > 9 Then
                MsgBox "Este fornecedor já possui 10 motoristas cadastrados. Exclua algum motorista para cadastrar o desejado.", vbInformation + vbOKOnly, "SisTrans"
                Exit Sub
            End If
            .Close
        End With
        
        .AddNew
    End If
        
    !nome_motorista = txt_NomeMotorista.Text
    !cnpj_fornec = txt_CNPJ.Text
    !OrgaoExped_Id = txt_OrgaoExp.Text
    !identidade = TXT_IDENTIDADE.Text
    !Dt_Val_Id = Txt_Validade.Text
    
    .UpdateBatch adAffectAll
    MsgBox "Arquivo salvo.", vbOKOnly + vbInformation, "SisTrans"
    .Close
    
    varCodMotorista = ""

End With

btn_Salvar.Enabled = True
btn_Excluir.Enabled = False

Call prc_LimparCampos

txt_NomeMotorista.SetFocus

Exit Sub
Erro:
    MsgBox "Todos os dados devem ser preenchidos.", vbOKOnly + vbInformation, "SisTrans"
    txt_NomeMotorista.SetFocus
End Sub
Private Sub Form_Load()
Call prc_LimparCampos
On Error Resume Next
Me.Top = 0
Me.Left = 0
txt_CNPJ.Text = varCNPJ
'1 - modo normal
'2 - modo adição
Select Case varModoAbrir
    Case 1
        btn_Salvar.Enabled = False
        btn_Excluir.Enabled = True
        Set rs_tabMotoristaFornec = New Recordset
        With rs_tabMotoristaFornec
            .Open "select * from tab_trans_motorista_fornecedor where cod_motorista =" & varCodMotorista & "", cnConexao, adOpenStatic, adLockOptimistic
            If .RecordCount = 1 Then
                TXT_IDENTIDADE.Text = !identidade
                txt_NomeMotorista.Text = !nome_motorista
                txt_CNPJ.Text = !cnpj_fornec
                txt_OrgaoExp.Text = !OrgaoExped_Id
                Txt_Validade.Text = !Dt_Val_Id
            End If
            .Close
        End With
    Case 2
        FRA_CADASTRO.Enabled = True
        btn_Salvar.Enabled = True
        btn_Excluir.Enabled = False
        txt_CNPJ.Text = varCNPJ
End Select
End Sub
Private Sub prc_LimparCampos()
If varModoAbrir <> 1 Then varCodMotorista = ""
txt_NomeMotorista.Text = ""
TXT_IDENTIDADE.Text = ""
txt_OrgaoExp.Text = ""
Txt_Validade.Mask = "          "
Txt_Validade.Mask = "99/99/9999"
End Sub
Private Sub txt_Identidade_LostFocus()
If TXT_IDENTIDADE = " " Then Exit Sub
var_SQL = "select * from tab_trans_Motorista_fornecedor "
var_SQL = var_SQL + " where identidade ='" & TXT_IDENTIDADE.Text & "'"
Set DS = New Recordset
DS.Open var_SQL, cnConexao, adOpenStatic, adLockOptimistic
With DS
     If .RecordCount = 1 Then
        txt_NomeMotorista = !nome_motorista
        txt_OrgaoExp = !OrgaoExped_Id
        Txt_Validade = !Dt_Val_Id
        btn_Excluir.Enabled = True
        Fra_Identidade.Enabled = False
        txt_NomeMotorista.Enabled = False
    Else
        Fra_Identidade.Enabled = True
        txt_NomeMotorista.Enabled = True
        
        btn_Salvar.Enabled = True
        btn_Excluir.Enabled = False
       
    End If
End With
End Sub
Private Sub txt_NomeFornec_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txt_NomeMotorista_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txt_OrgaoExp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
