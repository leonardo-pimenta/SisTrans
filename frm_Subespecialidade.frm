VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Subespecialidade 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de subespecialidades"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   6975
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Sal&var"
      Height          =   855
      Left            =   120
      Picture         =   "frm_Subespecialidade.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salvar"
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1080
      Picture         =   "frm_Subespecialidade.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Excluir"
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   6240
      Picture         =   "frm_Subespecialidade.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   2400
      Width           =   855
   End
   Begin VB.Frame FRA_CADASTRO 
      Height          =   1935
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6975
      Begin VB.TextBox TXT_ESPECIALIDADE 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid GRID_SUBESPECIALIDADE 
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   1296
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin VB.TextBox TXT_CPF 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DBCOM_SUBESPECIALIDADE 
         Height          =   315
         Left            =   3960
         TabIndex        =   2
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "SIGLA"
         BoundColumn     =   "SIGLA"
         Text            =   ""
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Especialidade.:"
         Height          =   195
         Left            =   1440
         TabIndex        =   11
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CPF.:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Subespecialidade:"
         Height          =   195
         Left            =   3960
         TabIndex        =   8
         Top             =   240
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frm_Subespecialidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExcluir_Click()
On Error GoTo Error

If DBCOM_SUBESPECIALIDADE = "" Then
   MsgBox "Selecione uma subespecialidade.", vbInformation, "SISTRANS"
   Exit Sub
End If
If MsgBox("Deseja excluir o registro selecionado?", vbYesNo + vbQuestion, "SisVTR") = vbYes Then
    cmdExcluir.Enabled = False
    cmdSalvar.Enabled = False
    SQL = "select * FROM TAB_GER_pessoa_subespecialidade"
    SQL = SQL + " WHERE sigla_subespecialidade='" & DBCOM_SUBESPECIALIDADE & "'"
    SQL = SQL + " and CPF_pessoa='" & TXT_CPF & "'"
    Set DS = New Recordset
    DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
    DS.Delete
End If
Call Form_Load
MsgBox "Registro deletado.", vbInformation, "SISTRANS"
cmdSalvar.Enabled = True
Exit Sub
Error:
    MsgBox "Erro. O registro não foi excluído,por fazer parte de um relacionamento.", vbOKOnly + vbInformation, "SisTrans"
    btn_Sair.SetFocus

End Sub
Private Sub cmdSair_Click()
Unload Me
frm_CadPessoa.TXT_CODINOME.SetFocus
End Sub
Private Sub cmdSalvar_Click()
If DBCOM_SUBESPECIALIDADE = "" Then
   MsgBox "Selecione uma subespecialidade.", vbInformation, "SISTRANS"
   Exit Sub
End If
SQL = "select * FROM Tab_Ger_pessoa_Subespecialidade"
SQL = SQL + " WHERE CPF_PESSOA='" & TXT_CPF & "'"
SQL = SQL + "AND SIGLA_SUBESPECIALIDADE='" & DBCOM_SUBESPECIALIDADE & "'"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 0 Then
   DS.AddNew
   DS!CPF_pessoa = TXT_CPF
   DS!SIGLA_subespecialidade = DBCOM_SUBESPECIALIDADE.Text
   DS.UpdateBatch adAffectAll

   SQL = "select * FROM Tab_Ger_pessoa_Subespecialidade"
   SQL = SQL + " WHERE CPF_PESSOA='" & TXT_CPF & "'"
   SQL = SQL + " order by sigla_subespecialidade"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   Set GRID_SUBESPECIALIDADE.DataSource = DS
   GRID_SUBESPECIALIDADE.Columns(0).Width = 2960
   GRID_SUBESPECIALIDADE.Columns(1).Width = 2960
   GRID_SUBESPECIALIDADE.Columns(0).Caption = "Subespecialidade"
   GRID_SUBESPECIALIDADE.Columns(1).Caption = "CPF"
   
   SQL = "select sigla_subespecialidade FROM Tab_Ger_pessoa_Subespecialidade"
   SQL = SQL + " WHERE CPF_PESSOA='" & TXT_CPF & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   Set frm_CadPessoa.DBCOM_SUBESPECIALIDADE.RowSource = DS
End If
End Sub
Private Sub Form_Load()
TXT_CPF = frm_CadPessoa.TXT_CPF
TXT_ESPECIALIDADE = frm_CadPessoa.COMBO_CORPO_QUADRO_ESPECIALIDADE.Text
'Combo Subespecialidades
SQL = "select sigla FROM Tab_Ger_Subespecialidade"
SQL = SQL + " order by sigla"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set DBCOM_SUBESPECIALIDADE.RowSource = DS
If frm_CadPessoa.TXT_CPF <> "" Then
   SQL = "select * FROM Tab_Ger_pessoa_Subespecialidade"
   SQL = SQL + " WHERE CPF_PESSOA='" & TXT_CPF & "'"
   SQL = SQL + " order by sigla_subespecialidade"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   Set GRID_SUBESPECIALIDADE.DataSource = DS
   GRID_SUBESPECIALIDADE.Columns(0).Width = 2960
   GRID_SUBESPECIALIDADE.Columns(1).Width = 2960
   GRID_SUBESPECIALIDADE.Columns(0).Caption = "Subespecialidade"
   GRID_SUBESPECIALIDADE.Columns(1).Caption = "CPF"
End If
End Sub
Private Sub GRID_SUBESPECIALIDADE_Click()
DBCOM_SUBESPECIALIDADE = GRID_SUBESPECIALIDADE.Columns(0)
End Sub
