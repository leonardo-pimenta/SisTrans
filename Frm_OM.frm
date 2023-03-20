VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frm_OM 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de OM"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7170
   ControlBox      =   0   'False
   Icon            =   "Frm_OM.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btn_Cartao 
      Caption         =   "Veí&culo"
      Height          =   855
      Left            =   3360
      Picture         =   "Frm_OM.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Editar"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Btn_Cartao_OM 
      Caption         =   "C&artão de OM"
      Enabled         =   0   'False
      Height          =   855
      Left            =   4560
      Picture         =   "Frm_OM.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Editar"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   6240
      Picture         =   "Frm_OM.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "E&ditar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   1080
      Picture         =   "Frm_OM.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Editar"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   855
      Left            =   2040
      Picture         =   "Frm_OM.frx":1412
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Excluir"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Sal&var"
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      Picture         =   "Frm_OM.frx":1854
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salvar"
      Top             =   3960
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   6975
   End
   Begin VB.Frame FRA_CADASTRO 
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   6975
      Begin VB.TextBox TXT_IND_NAV 
         Height          =   285
         Left            =   4440
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1080
         Width           =   735
      End
      Begin MSMask.MaskEdBox TXT_COD_OM 
         Height          =   300
         Left            =   360
         TabIndex        =   0
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   5
         Mask            =   "99999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TXT_NOME 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TXT_SIGLA 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   30
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Indicativo:"
         Height          =   195
         Left            =   4440
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sigla:"
         Height          =   195
         Left            =   1440
         TabIndex        =   14
         Top             =   840
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Código.:"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   585
      End
   End
   Begin MSDataGridLib.DataGrid GRID_OM 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   0   'False
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
Attribute VB_Name = "Frm_OM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Cartao_Click()
vgl_Cor = "Lilás"
frm_CadVeiculo.vgl_Responsavel = 3
frm_CadVeiculo.opt_OM.Value = True
frm_CadVeiculo.txt_Responsavel = TXT_COD_OM.Text
frm_CadVeiculo.txt_Descricao = TXT_NOME.Text
End Sub
Private Sub Btn_Cartao_OM_Click()
If TXT_COD_OM = "" Then
   MsgBox "Digite o código da OM", vbInformation + vbOKOnly, "SisTrans"
Else
   SQL = "select * FROM TAB_GER_OM"
   SQL = SQL + " where COD='" & TXT_COD_OM & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   If DS.RecordCount <> 0 Then
      Var_COD = DS!cod
      Var_nome = DS!NOME
      vgl_TipoResponsavel = "OM"
      frm_Cartao_Emitir.txt_Codigo = Var_COD
      frm_Cartao_Emitir.txt_Descricao = Var_nome
      vgl_Cor = "Azul"
      frm_Cartao_Emitir.Txt_Tarja = "AZUL"
   End If
End If
End Sub
Private Sub cmdEditar_Click()
VarAcao = "Registro Velho"
cmdSalvar.Enabled = True
cmdEditar.Enabled = False
cmdExcluir.Enabled = True
fra_Cadastro.Enabled = True
TXT_COD_OM.Enabled = False
TXT_NOME.SetFocus
End Sub
Private Sub cmdExcluir_Click()
On Error GoTo Error
If MsgBox("Deseja excluir o registro selecionado?", vbYesNo + vbQuestion, "SisVTR") = vbYes Then
    cmdEditar.Enabled = False
    cmdExcluir.Enabled = False
    cmdSalvar.Enabled = False
    SQL = "select * FROM TAB_GER_OM"
    SQL = SQL + " WHERE COD='" & TXT_COD_OM & "'"
    Set DS = New Recordset
    DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
    DS.Delete
End If
MsgBox "Registro deletado.", vbInformation, "SISTRANS"
cmdSalvar.Enabled = True
Call Form_Load
TXT_COD_OM.Text = "_____"
TXT_COD_OM.Mask = "99999"
Exit Sub
Error:
    MsgBox "Erro. O registro não foi excluído,por fazer parte de um relacionamento.", vbOKOnly + vbInformation, "SisTrans"
    cmdSair.SetFocus
End Sub
Private Sub cmdSair_Click()
Unload Me
End Sub
Private Sub cmdSalvar_Click()
If TXT_COD_OM = "" Then
   MsgBox "Digite o código da OM.", vbInformation, "SISTRANS"
   Exit Sub
End If
If TXT_SIGLA = "" Then
   MsgBox "Digite a sigla da OM.", vbInformation, "SISTRANS"
   Exit Sub
End If
If TXT_NOME = "" Then
   MsgBox "Digite o nome da OM.", vbInformation, "SISTRANS"
   Exit Sub
End If
If TXT_IND_NAV = "" Then
   MsgBox "Digite o Indicativo naval da OM.", vbInformation, "SISTRANS"
   Exit Sub
End If
Call Salvar
Call Form_Load
TXT_COD_OM.Text = "_____"
TXT_COD_OM.Mask = "99999"
End Sub
Private Sub Salvar()
On Error GoTo Error
SQL = "select * FROM TAB_GER_OM"
SQL = SQL + " where COD='" & TXT_COD_OM & "'"
Set ds1 = New Recordset
ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
If VarAcao = "Registro Novo" Then
   ds1.AddNew
   ds1!cod = TXT_COD_OM
   ds1!SIGLA = TXT_SIGLA
   ds1!NOME = TXT_NOME
   ds1!IND_NAV = TXT_IND_NAV
   ds1.UpdateBatch adAffectAll
   MsgBox "Registro salvo.", vbInformation, "SisVTR"
ElseIf VarAcao = "Registro Velho" Then
   ds1!cod = TXT_COD_OM
   ds1!SIGLA = TXT_SIGLA
   ds1!NOME = TXT_NOME
   ds1!IND_NAV = TXT_IND_NAV
   ds1.UpdateBatch adAffectAll
   MsgBox "Registro alterado.", vbInformation, "SisVTR"
End If
TXT_COD_OM.Enabled = True
Exit Sub
Error:
    MsgBox "Erro. O registro não pode ser alterado,por fazer parte de um relacionamento.", vbOKOnly + vbInformation, "SisTrans"
    TXT_COD_OM.Enabled = True
End Sub
Private Sub Form_Load()
Frm_OM.Top = 0
Frm_OM.Left = 0
Call LIMPAR
SQL = "select * FROM TAB_GER_OM"
SQL = SQL + " ORDER BY SIGLA"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set GRID_OM.DataSource = DS
GRID_OM.Columns(0).Width = 1000
GRID_OM.Columns(1).Width = 1000
GRID_OM.Columns(2).Width = 3610
GRID_OM.Columns(3).Width = 1000
fra_Cadastro.Enabled = True
cmdSalvar.Enabled = False
cmdEditar.Enabled = False
cmdExcluir.Enabled = False
btn_Cartao.Enabled = False
Btn_Cartao_OM.Enabled = False
End Sub
Private Sub GRID_OM_dblClick()
TXT_COD_OM = GRID_OM.Columns(0)
TXT_NOME = GRID_OM.Columns(2)
TXT_SIGLA = GRID_OM.Columns(1)
TXT_IND_NAV = GRID_OM.Columns(3)
Call TXT_COD_OM_LostFocus
End Sub
Private Sub TXT_COD_OM_Change()
SQL = "select * FROM TAB_GER_OM"
SQL = SQL + " where cod LIKE '" & "%" & "' + '" & TXT_COD_OM & "' + '" & "%" & "' ORDER BY cod"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

Set GRID_OM.DataSource = DS
GRID_OM.Columns(0).Width = 1000
GRID_OM.Columns(1).Width = 1000
GRID_OM.Columns(2).Width = 3610
GRID_OM.Columns(3).Width = 1000
End Sub
Private Sub TXT_COD_OM_GotFocus()
Call LIMPAR
TXT_COD_OM.Text = "_____"
TXT_COD_OM.Mask = "99999"
End Sub
Private Sub TXT_COD_OM_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_IND_NAV_Change()
SQL = "select * FROM TAB_GER_OM"
SQL = SQL + " where IND_NAV LIKE '" & "%" & "' + '" & TXT_IND_NAV & "' + '" & "%" & "' ORDER BY IND_NAV"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set GRID_OM.DataSource = DS
GRID_OM.Columns(0).Width = 1000
GRID_OM.Columns(1).Width = 1000
GRID_OM.Columns(2).Width = 3610
GRID_OM.Columns(3).Width = 1000
End Sub
Private Sub TXT_IND_NAV_LostFocus()
If TXT_IND_NAV.Text = "" Then
   Exit Sub
ElseIf VarAcao <> "Registro Velho" Then
   Call TXT_COD_OM_LostFocus
End If
End Sub
Private Sub TXT_NOME_Change()

SQL = "select * FROM TAB_GER_OM"
SQL = SQL + " where nome LIKE '" & "%" & "' + '" & TXT_NOME & "' + '" & "%" & "' ORDER BY nome"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set GRID_OM.DataSource = DS
GRID_OM.Columns(0).Width = 1000
GRID_OM.Columns(1).Width = 1000
GRID_OM.Columns(2).Width = 3610
GRID_OM.Columns(3).Width = 1000
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_NOME_LostFocus()
If TXT_NOME.Text = "" Then
   Exit Sub
ElseIf VarAcao <> "Registro Velho" Then
   Call TXT_COD_OM_LostFocus
End If
End Sub
Private Sub TXT_SIGLA_Change()
SQL = "select * FROM TAB_GER_OM"
SQL = SQL + " where SIGLA LIKE '" & "%" & "' + '" & TXT_SIGLA & "' + '" & "%" & "' ORDER BY sigla"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set GRID_OM.DataSource = DS
GRID_OM.Columns(0).Width = 1000
GRID_OM.Columns(1).Width = 1000
GRID_OM.Columns(2).Width = 3610
GRID_OM.Columns(3).Width = 1000
End Sub
Private Sub TXT_SIGLA_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_IND_NAV_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_COD_OM_LostFocus()
If TXT_COD_OM.Text = "_____" Or TXT_COD_OM.Text = "" Then
   Exit Sub
Else
   SQL = "select * FROM TAB_GER_OM"
   SQL = SQL + " where COD='" & TXT_COD_OM & "'"
   Set ds1 = New Recordset
   ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   If ds1.RecordCount = 1 Then
      VarAcao = "Registro Velho"
      cmdSalvar.Enabled = False
      cmdEditar.Enabled = True
      cmdExcluir.Enabled = True
      btn_Cartao.Enabled = True
      Btn_Cartao_OM.Enabled = True
      
      
      TXT_COD_OM = ds1!cod
      TXT_SIGLA = ds1!SIGLA
      TXT_NOME = ds1!NOME
      TXT_IND_NAV = ds1!IND_NAV
   Else
      VarAcao = "Registro Novo"
      cmdSalvar.Enabled = True
      cmdEditar.Enabled = False
      cmdExcluir.Enabled = False
   End If
End If
End Sub
Private Sub LIMPAR()
TXT_SIGLA = ""
TXT_NOME = ""
TXT_IND_NAV = ""
End Sub
Private Sub TXT_SIGLA_LostFocus()
If TXT_SIGLA.Text = "" Then
   Exit Sub
ElseIf VarAcao <> "Registro Velho" Then
   Call TXT_COD_OM_LostFocus
End If
End Sub
