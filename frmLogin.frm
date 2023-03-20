VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Login 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SisTrans - Login de acesso"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_Sobre 
      Caption         =   "&Sobre"
      Height          =   780
      Left            =   120
      Picture         =   "frmLogin.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Efetua o login no sistema"
      Top             =   1440
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   5775
   End
   Begin VB.TextBox txt_Senha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Entre com a senha correspondente"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txt_Usuario 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   0
      ToolTipText     =   "Entre com a senha correspondente"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   780
      Left            =   5160
      Picture         =   "frmLogin.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancela o login e sai do sistema"
      Top             =   1440
      Width           =   750
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   780
      Left            =   4320
      Picture         =   "frmLogin.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Efetua o login no sistema"
      Top             =   1440
      Width           =   750
   End
   Begin VB.TextBox txt_Descricao 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Entre com a senha correspondente"
      Top             =   480
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog cdg_Conexao 
      Left            =   5400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Senha:"
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   840
      Width           =   510
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   1080
      TabIndex        =   7
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Usuário:"
      Height          =   195
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   585
   End
   Begin VB.Image Image3 
      Height          =   960
      Left            =   120
      Picture         =   "frmLogin.frx":1108
      Stretch         =   -1  'True
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frm_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCadSenha As Recordset

Dim OpUsuario As String
Dim opLogin As String

Private Sub btn_Sobre_Click()
frm_Sobre.Show vbModal
End Sub
Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdLogin_Click()

Call Login

End Sub

Private Sub txt_Senha_GotFocus()

Set rsCadSenha = New Recordset

With rsCadSenha
    .Open "select * from tab_trans_aux_senha where login ='" & txt_Usuario.Text & "'", cnConexao, adOpenStatic, adLockOptimistic

    If .RecordCount = 0 Then
        MsgBox "Usuário inválido.", vbOKOnly + vbInformation, "Login de Acesso"
        
        .Close
        
        txt_Descricao.Text = ""
        txt_Usuario.Text = ""
        txt_Usuario.SetFocus
    Else
        txt_Descricao.Text = !descricao
    End If

End With

End Sub

Private Sub txt_Usuario_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()

'Mostra automaticamente a versão do sistema
frm_Login.Caption = "SisTrans - Versão " & App.Major & "." & App.Minor & "." & App.Revision

Call OpenConexao("", "SisTrans")

End Sub

Private Sub tmrlblSenha_Timer()

If lblSenha.Visible = True Then
    lblSenha.Visible = False
    Exit Sub
ElseIf lblSenha.Visible = False Then
    lblSenha.Visible = True
    Exit Sub
End If

End Sub

Public Sub Login()

Dim var_Nivel As String

If txt_Descricao.Text = "" Then Exit Sub
  
If txt_Senha.Text = rsCadSenha!senha Then

        vgl_Responsavel = txt_Usuario
        vgl_Nivel = rsCadSenha!nivel
        
        rsCadSenha.Close
        Set rsCadSenha = Nothing
        
        Unload frm_Login
        frm_Principal.Show
       
    Exit Sub
Else
    MsgBox "Senha inválida sr " & txt_Usuario & " !", , "Login de Acesso"
    txt_Senha.Text = ""
    txt_Senha.SetFocus
End If


End Sub

Private Sub txt_Senha_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

