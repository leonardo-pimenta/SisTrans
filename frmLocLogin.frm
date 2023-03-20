VERSION 5.00
Begin VB.Form frmLocLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Configuração de senhas"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3435
   Icon            =   "frmLocLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txt2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txt1 
         Height          =   285
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   8
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lbl2 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lbl1 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Usuário:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   105
         TabIndex        =   6
         Top             =   120
         Width           =   1800
      End
   End
   Begin VB.Frame FRA2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   735
      Left            =   2160
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
      Begin VB.OptionButton optUser 
         BackColor       =   &H00404000&
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptAdmin 
         BackColor       =   &H00404000&
         Caption         =   "Admin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame fra3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1215
      Left            =   2160
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "&Confirmar"
         Height          =   855
         Left            =   120
         Picture         =   "frmLocLogin.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Image imgArm 
      Height          =   2565
      Index           =   1
      Left            =   600
      Picture         =   "frmLocLogin.frx":0884
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblOperacao 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      TabIndex        =   10
      Top             =   3480
      Width           =   3480
   End
   Begin VB.Menu mnuadicioanr 
      Caption         =   "&Adicionar"
   End
   Begin VB.Menu mnueditar 
      Caption         =   "&Editar"
   End
   Begin VB.Menu mnuexcluir 
      Caption         =   "E&xcluir"
   End
   Begin VB.Menu mnusair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "frmLocLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCadSenha As Recordset
Dim xOperacao As Byte
Dim xUsuario As String
Dim xSenha As String
Dim xNSenha As String
Dim xCSenha As String
Dim xOP As String

Private Sub cmdSalvar_Click()

Dim count As Integer
Dim xCount As Integer


If xOperacao = 3 Then

    Set rsCadSenha = New Recordset
    
    'Call OpenConeccao
    
    With rsCadSenha
    .Open "select * from cadsenha order by usuario", cnConexao, adOpenStatic, adLockOptimistic
            
    count = .RecordCount
        
    If MsgBox("Deseja excluir?", vbYesNo + vbQuestion, " Excluir") = vbYes Then
        For xCount = 1 To count
            If !usuario = xUsuario And !senha = xSenha Then
                .Delete
                MsgBox "Usuario excluído êxito.", , "Excluir"
                .Close
            
                Call Iniciar
                Exit Sub
            End If
        
            .MoveNext
        Next
                            
    End If
    
    MsgBox "A senha não confere.", , "Operação cancelada"
    
    Call Iniciar
    Exit Sub
    
    .Close
    
    Call Iniciar
    
    End With
    
End If

If MsgBox("Deseja salvar?", vbYesNo + vbQuestion, "Adicionar") = vbYes Then
    
    If OptAdmin.Value = True Then xOP = "ADMIN"
    If optUser.Value = True Then xOP = "USER"
    
    'Call OpenConeccao
    
    Set rsCadSenha = New Recordset
    
    With rsCadSenha
        .Open "select * from cadsenha order by usuario", cnConexao, adOpenStatic, adLockOptimistic
            
        count = .RecordCount
                
        
        
            If xOperacao = 1 Then
                .AddNew
                    !usuario = xUsuario
                    !senha = xSenha
                    !op = xOP
                .UpdateBatch adAffectAll
                MsgBox "Senha salvo com êxito.", , "Adicionar"
                .Close
            
                Call Iniciar
                Exit Sub
            
            End If
            
            If xOperacao = 2 Then
                For xCount = 1 To count
                    If !usuario = xUsuario And !senha = xSenha Then
                            !usuario = xUsuario
                            !senha = xCSenha
                        .UpdateBatch adAffectAll
                        MsgBox "Senha salvo com êxito.", , "Editar"
                    .Close
                    
                    Call Iniciar
                    Exit Sub
                    
                    End If
                                    
                    .MoveNext
                Next
                
                MsgBox "A antiga senha não confere.", , "Operação cancelada"
                
                Call Iniciar
                Exit Sub
                
            End If
            
    End With
    
End If

Call Iniciar

End Sub

Private Sub mnuadicioanr_Click()
xOperacao = 1

fra1.Enabled = True
FRA2.Enabled = True
fra3.Enabled = True

txt1.SetFocus

lblOperacao.Caption = "Adicionar"

End Sub

Private Sub mnueditar_Click()
xOperacao = 2

fra1.Enabled = True
fra3.Enabled = True

txt1.SetFocus

lblOperacao.Caption = "Editar"

End Sub

Private Sub mnuexcluir_Click()
xOperacao = 3

fra1.Enabled = True
fra3.Enabled = True

txt1.SetFocus

lblOperacao.Caption = "Excluir"

End Sub

Private Sub mnusair_Click()
Unload Me
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt1_LostFocus()
txt1.Enabled = False
txt1.PasswordChar = "*"
End Sub

Private Sub txt2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt2_LostFocus()

If xOperacao = 1 Then
    
    If lbl2.Caption = "Confirme a Senha:" Then
        xCSenha = txt2.Text
        If xSenha <> xCSenha Then
            MsgBox "Senhas não conferem", , "Adicionar"
            txt2.Text = ""
            txt2.SetFocus
            Exit Sub
        End If
        fra1.Enabled = False
        OptAdmin.SetFocus
        Exit Sub
    End If
    
    xUsuario = txt1.Text
    xSenha = txt2.Text
    lbl2.Caption = "Confirme a Senha:"
    txt2.Text = ""
    txt2.SetFocus

End If
        
If xOperacao = 2 Then
    
    If lbl1.Caption = "Nova Senha:" Then
        xNSenha = txt1.Text
        xCSenha = txt2.Text
        If xCSenha <> xNSenha Then
            MsgBox "Senhas não conferem", , "Adicionar"
            txt1.Text = ""
            txt2.Text = ""
            txt1.Enabled = True
            txt1.SetFocus
            Exit Sub
        End If
        fra1.Enabled = False
        cmdSalvar.SetFocus
        Exit Sub
    End If
    
    xUsuario = txt1.Text
    xSenha = txt2.Text
    lbl1.Caption = "Nova Senha:"
    lbl2.Caption = "Confirme a Senha:"
    txt1.Text = ""
    txt2.Text = ""
    txt1.Enabled = True
    txt1.SetFocus

End If


If xOperacao = 3 Then
    fra1.Enabled = False
    xUsuario = txt1.Text
    xSenha = txt2.Text
    cmdSalvar.SetFocus
    Exit Sub
End If

End Sub

Private Sub Iniciar()

fra1.Enabled = False
FRA2.Enabled = False
fra3.Enabled = False

lbl1.Caption = "Usuário:"
lbl2.Caption = "Senha:"

txt1.Text = ""
txt2.Text = ""

txt1.Enabled = True
txt2.Enabled = True

txt1.PasswordChar = ""

'Call CloseConeccao

End Sub
