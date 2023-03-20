VERSION 5.00
Begin VB.Form frm_Senha_Alterar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alteração de senha"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CommandButton btn_Sair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   855
      Left            =   2640
      Picture         =   "frm_Senha_Alterar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton btn_Salvar 
      Caption         =   "S&alvar"
      Default         =   -1  'True
      Height          =   855
      Left            =   120
      Picture         =   "frm_Senha_Alterar.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salvar"
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txt_ConfirSenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txt_NovaSenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txt_SenhaAtual 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgArm 
      Height          =   525
      Index           =   1
      Left            =   360
      Picture         =   "frm_Senha_Alterar.frx":074C
      Stretch         =   -1  'True
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Confirmar senha:"
      Height          =   195
      Left            =   1320
      TabIndex        =   7
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nova senha:"
      Height          =   195
      Left            =   1560
      TabIndex        =   6
      Top             =   480
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Senha atual:"
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frm_Senha_Alterar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_tabSenha As Recordset

Private Sub btn_Sair_Click()
Unload Me
End Sub

Private Sub btn_Salvar_Click()

Dim varSenhaAtual As String
Dim vCount As Byte

'Verifica se as novas senhas conferem
If txt_NovaSenha.Text <> txt_ConfirSenha.Text Then
    MsgBox "Nova senha não confere.", vbOKOnly, "SisTrans"
    
    txt_NovaSenha.Text = ""
    txt_ConfirSenha.Text = ""
    
    txt_NovaSenha.SetFocus
    Exit Sub
End If

'Abre a tabela verifica se a senha esta correta, caso sim edita-a
Set rs_tabSenha = New Recordset

With rs_tabSenha
    .Open "select * from tab_trans_aux_senha where login ='" & vgl_Responsavel & "' and senha ='" & txt_SenhaAtual.Text & "'", cnConexao, adOpenStatic, adLockOptimistic

    If .RecordCount = 0 Then
        MsgBox "Senha atual incorreta.", vbOKOnly + vbInformation, "SisTrans"
            
        txt_SenhaAtual.Text = ""
        txt_SenhaAtual.SetFocus
    Else
        !senha = txt_NovaSenha
        .UpdateBatch adAffectAll
        MsgBox "Senha alterada.", vbOKOnly, "SisTrans"
        .Close
        
        Unload Me
        
    End If

End With

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Caption = "Alteração de senha - User: " & vgl_Responsavel

Me.Left = vgl_X
Me.Top = vgl_Y

End Sub

Private Sub txt_ConfirSenha_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_NovaSenha_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_SenhaAtual_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
