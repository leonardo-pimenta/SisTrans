VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Especialidade_Subespecialidade 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Especialidades e Subespecialidades"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid GRID_ESPECIALIDADE 
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2778
      _Version        =   393216
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame FRA_SUBESPECIALIDADE 
      Caption         =   "Subespecialidade"
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   6975
      Begin VB.TextBox TXT_SUB_DESCRICAO 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox TXT_SUB_SIGLA 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descrição.:"
         Height          =   195
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sigla:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.Frame FRA_ESPECIALIDADE 
      Caption         =   "Especialidade"
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   6975
      Begin VB.TextBox TXT_ESPEC_SIGLA 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   0
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TXT_ESPEC_DESCRICAO 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Sigla:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição.:"
         Height          =   195
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   6240
      Picture         =   "frm_Especialidade_Subespecialidade.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "E&ditar"
      Height          =   855
      Left            =   1080
      Picture         =   "frm_Especialidade_Subespecialidade.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Editar"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   2040
      Picture         =   "frm_Especialidade_Subespecialidade.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Excluir"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Sal&var"
      Height          =   855
      Left            =   120
      Picture         =   "frm_Especialidade_Subespecialidade.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salvar"
      Top             =   4200
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   6975
   End
   Begin MSDataGridLib.DataGrid GRID_SUBESPECIALIDADE 
      Height          =   1575
      Left            =   3600
      TabIndex        =   5
      Top             =   2280
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2778
      _Version        =   393216
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_Especialidade_Subespecialidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEditar_Click()
cmdSalvar.Enabled = True
cmdEditar.Enabled = False
cmdExcluir.Enabled = True
If TXT_ESPEC_DESCRICAO.Text <> "" Then
   TXT_ESPEC_DESCRICAO.Enabled = True
   TXT_ESPEC_DESCRICAO.SetFocus
End If
If TXT_SUB_DESCRICAO.Text <> "" Then
   TXT_SUB_DESCRICAO.Enabled = True
   TXT_SUB_DESCRICAO.SetFocus
End If
End Sub
Private Sub cmdExcluir_Click()
On Error GoTo Error
If TXT_ESPEC_SIGLA <> "" Then
   SQL = "select SIGLA FROM TAB_GER_ESPECIALIDADE"
   SQL = SQL + " WHERE SIGLA='" & TXT_ESPEC_SIGLA & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   If MsgBox("Deseja excluir o registro selecionado?", vbYesNo + vbQuestion, "SisVTR") = vbYes Then
      cmdEditar.Enabled = False
      cmdExcluir.Enabled = False
      DS.Delete
      DS.UpdateBatch adAffectAll
      Call LIMPAR
      Call Form_Load
      MsgBox "Registro deletado.", vbInformation, "SisVTR"
   End If
End If
If TXT_SUB_SIGLA <> "" Then
   SQL = "select SIGLA FROM TAB_GER_SUBESPECIALIDADE"
   SQL = SQL + " WHERE SIGLA='" & TXT_SUB_SIGLA & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   If MsgBox("Deseja excluir o registro selecionado?", vbYesNo + vbQuestion, "SisVTR") = vbYes Then
      cmdEditar.Enabled = False
      cmdExcluir.Enabled = False
      DS.Delete
      DS.UpdateBatch adAffectAll
      Call LIMPAR
      Call Form_Load
      MsgBox "Registro deletado.", vbInformation, "SisVTR"
   End If
End If
Exit Sub
Error:
    MsgBox "Erro. O registro não foi excluído,por fazer parte de um relacionamento.", vbOKOnly + vbInformation, "SisTrans"
    cmdSair.SetFocus
End Sub
Private Sub cmdSair_Click()
Unload Me
End Sub
Private Sub cmdSalvar_Click()
If TXT_ESPEC_SIGLA <> "" And TXT_ESPEC_DESCRICAO <> "" Then
    SQL = "select * FROM TAB_GER_ESPECIALIDADE"
    SQL = SQL + " WHERE SIGLA='" & TXT_ESPEC_SIGLA & "'"
    Set DS = New Recordset
    DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
    If DS.RecordCount = 0 Then
       DS.AddNew
    End If
      DS!SIGLA = TXT_ESPEC_SIGLA
      DS!Desc = TXT_ESPEC_DESCRICAO
      DS.UpdateBatch adAffectAll

    If VarAcao_ESPEC = "Registro Novo" Then
       MsgBox "Registro salvo.", vbInformation, "SisVTR"
    ElseIf VarAcao_ESPEC = "Registro Velho" Then
       MsgBox "Registro alterado.", vbInformation, "SisVTR"
    End If
End If
If TXT_SUB_SIGLA <> "" And TXT_SUB_DESCRICAO <> "" Then
   SQL = "select * FROM TAB_GER_SUBESPECIALIDADE"
   SQL = SQL + " WHERE SIGLA='" & TXT_SUB_SIGLA & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   If DS.RecordCount = 0 Then
      DS.AddNew
   End If
      DS!SIGLA = TXT_SUB_SIGLA
      DS!Desc = TXT_SUB_DESCRICAO
      DS.UpdateBatch adAffectAll
   If VarAcao_SUB = "Registro Novo" Then
      MsgBox "Registro salvo.", vbInformation, "SisVTR"
   Else
      MsgBox "Registro alterado.", vbInformation, "SisVTR"
   End If
End If
Call LIMPAR
Call Form_Load
End Sub
Private Sub LIMPAR()
TXT_SUB_SIGLA = ""
TXT_SUB_DESCRICAO = ""
TXT_ESPEC_SIGLA = ""
TXT_ESPEC_DESCRICAO = ""
End Sub
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
'GRID_ESPECIALIDADE
SQL = "select * FROM TAB_GER_ESPECIALIDADE"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set GRID_ESPECIALIDADE.DataSource = DS
GRID_ESPECIALIDADE.Columns(1).Width = 1000
GRID_ESPECIALIDADE.Columns(2).Width = 3000
GRID_ESPECIALIDADE.Columns(1).Caption = "Espec.-Sigla "
GRID_ESPECIALIDADE.Columns(2).Caption = "Espec.-Descrição"
GRID_ESPECIALIDADE.Columns(0).Visible = False

'GRID_SUBESPECIALIDADE
SQL = "select * FROM TAB_GER_SUBESPECIALIDADE"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set GRID_SUBESPECIALIDADE.DataSource = DS
GRID_SUBESPECIALIDADE.Columns(0).Width = 1300
GRID_SUBESPECIALIDADE.Columns(1).Width = 3000
GRID_SUBESPECIALIDADE.Columns(0).Caption = "Subespec.-Sigla"
GRID_SUBESPECIALIDADE.Columns(1).Caption = "Subespec.-Descrição"
End Sub
Private Sub GRID_ESPECIALIDADE_dblClick()
TXT_ESPEC_SIGLA = GRID_ESPECIALIDADE.Columns(1)
TXT_ESPEC_DESCRICAO = GRID_ESPECIALIDADE.Columns(2)
Call TXT_ESPEC_SIGLA_LostFocus
End Sub
Private Sub GRID_SUBESPECIALIDADE_dblClick()
TXT_SUB_SIGLA = GRID_SUBESPECIALIDADE.Columns(0)
TXT_SUB_DESCRICAO = GRID_SUBESPECIALIDADE.Columns(1)
Call TXT_SUB_SIGLA_LostFocus
End Sub
Private Sub TXT_ESPEC_DESCRICAO_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_SUB_DESCRICAO_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_SUB_SIGLA_CHANGE()
TXT_SUB_DESCRICAO = ""
End Sub
Private Sub TXT_SUB_SIGLA_CLICK()
TXT_SUB_DESCRICAO = ""
End Sub
Private Sub TXT_ESPEC_SIGLA_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_SUB_SIGLA_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_ESPEC_SIGLA_CHANGE()
TXT_ESPEC_DESCRICAO = ""
End Sub
Private Sub TXT_ESPEC_SIGLA_CLICK()
TXT_ESPEC_DESCRICAO = ""
End Sub
Private Sub TXT_SUB_SIGLA_LostFocus()
If TXT_SUB_SIGLA <> "" Then
   SQL = "select * FROM TAB_GER_SUBESPECIALIDADE"
   SQL = SQL + " where SIGLA='" & TXT_SUB_SIGLA & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   If DS.RecordCount = 1 Then
      VarAcao_SUB = "Registro Velho"
    
      cmdSalvar.Enabled = False
      cmdEditar.Enabled = True
      cmdExcluir.Enabled = True
      TXT_SUB_DESCRICAO.Enabled = False
        
      TXT_SUB_DESCRICAO = DS!Desc
   Else
      VarAcao_SUB = "Registro Novo"
      cmdSalvar.Enabled = True
      cmdEditar.Enabled = False
      cmdExcluir.Enabled = False
      TXT_SUB_DESCRICAO.Enabled = True
   End If
End If
End Sub
Private Sub TXT_ESPEC_SIGLA_LostFocus()
If TXT_ESPEC_SIGLA <> "" Then
   SQL = "select * FROM TAB_GER_ESPECIALIDADE"
   SQL = SQL + " where SIGLA='" & TXT_ESPEC_SIGLA & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   If DS.RecordCount = 1 Then
      VarAcao_ESPEC = "Registro Velho"
   
      cmdSalvar.Enabled = False
      cmdEditar.Enabled = True
      cmdExcluir.Enabled = True
      TXT_ESPEC_DESCRICAO.Enabled = False
        
      TXT_ESPEC_DESCRICAO = DS!Desc
   Else
      VarAcao_ESPEC = "Registro Novo"
      cmdSalvar.Enabled = True
      cmdEditar.Enabled = False
      cmdExcluir.Enabled = False
      TXT_ESPEC_DESCRICAO.Enabled = True
   End If
End If
End Sub
Private Sub LIMPAR_SUBESPECIALIDADE()
TXT_SUB_SIGLA = ""
TXT_SUB_DESCRICAO = ""
End Sub
Private Sub LIMPAR_ESPECIALIDADE()
TXT_ESPEC_SIGLA = ""
TXT_ESPEC_DESCRICAO = ""
End Sub
