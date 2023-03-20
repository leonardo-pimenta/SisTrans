VERSION 5.00
Begin VB.Form frm_Senha_Incluir 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Novo usuário"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmb_Nivel 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      ItemData        =   "frm_Senha_Incluir.frx":0000
      Left            =   3240
      List            =   "frm_Senha_Incluir.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   4455
   End
   Begin VB.TextBox txt_ConfirSenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3600
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   4455
   End
   Begin VB.CommandButton btn_Sair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   855
      Left            =   3720
      Picture         =   "frm_Senha_Incluir.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton btn_Salvar 
      Caption         =   "S&alvar"
      Default         =   -1  'True
      Height          =   855
      Left            =   120
      Picture         =   "frm_Senha_Incluir.frx":030E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salvar"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txt_Senha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3600
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txt_Login 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   720
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txt_Descricao 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nível:"
      Height          =   195
      Left            =   2745
      TabIndex        =   12
      Top             =   120
      Width           =   435
   End
   Begin VB.Image imgArm 
      Height          =   525
      Index           =   1
      Left            =   360
      Picture         =   "frm_Senha_Incluir.frx":0750
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Confirma senha:"
      Height          =   195
      Left            =   2400
      TabIndex        =   11
      Top             =   1680
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Senha:"
      Height          =   195
      Left            =   3000
      TabIndex        =   9
      Top             =   1320
      Width           =   510
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Login:"
      Height          =   195
      Left            =   150
      TabIndex        =   8
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descrição do usuário:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1545
   End
End
Attribute VB_Name = "frm_Senha_Incluir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_tabSenha As Recordset

Private Sub btn_Sair_Click()
Unload Me
End Sub


Private Sub btn_Salvar_Click()

Dim vCount As Byte



If txt_Login.Text = "" Then
    MsgBox "É necessário entrar com um login.", vbOKOnly + vbInformation, "SisTrans"
    
    txt_Login.SetFocus
    Exit Sub
End If

'Verifica se as senhas conferem
If txt_Senha.Text <> txt_ConfirSenha.Text Then
    MsgBox "Senha não confere.", vbOKOnly, "SisTrans"
    
    txt_Senha.Text = ""
    txt_ConfirSenha.Text = ""
    
    txt_Senha.SetFocus
    Exit Sub
End If

'Abre a tabela verifica se a senha esta correta, caso sim edita-a
Set rs_tabSenha = New Recordset

With rs_tabSenha

'''''
    .Open "select * from tab_trans_aux_senha where login ='" & txt_Login.Text & "'", cnConexao, adOpenStatic, adLockOptimistic

    If .RecordCount = 0 Then
        .AddNew
            !Login = txt_Login.Text
            !descricao = txt_Descricao.Text
            !senha = txt_Senha.Text
            !nivel = cmb_Nivel.Text
        .UpdateBatch adAffectAll
        MsgBox "Usuário cadastrado.", vbOKOnly, "SisTrans"
        .Close
        
        Unload Me
    Else
        MsgBox "Este usuário já existe.", vbOKOnly, "SisTrans"
        
        txt_Senha.Text = ""
        txt_ConfirSenha.Text = ""
        txt_Login.Text = ""
        txt_Login.SetFocus
    End If

End With

End Sub
Private Sub Form_Load()

Me.Top = 0
Me.Left = 0

If vgl_Nivel = "ADMIN" Then
    'Possui todos os direitos dentro do sistema -Usuário da DN101
    
    frm_Senha_Incluir.cmb_Nivel.Clear
    frm_Senha_Incluir.cmb_Nivel.AddItem ("ADMIN")
    frm_Senha_Incluir.cmb_Nivel.AddItem ("SUPERVISOR")
    frm_Senha_Incluir.cmb_Nivel.AddItem ("USER")
    frm_Senha_Incluir.cmb_Nivel.AddItem ("CGS")


ElseIf vgl_Nivel = "SUPERVISOR" Then
    
    frm_Senha_Incluir.cmb_Nivel.Clear
    frm_Senha_Incluir.cmb_Nivel.AddItem ("SUPERVISOR")
    frm_Senha_Incluir.cmb_Nivel.AddItem ("USER")
    frm_Senha_Incluir.cmb_Nivel.AddItem ("CGS")
    
End If

End Sub
Private Sub txt_ConfirSenha_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_Descricao_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_Login_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_Senha_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
