VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_PostoGrad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SisVTR - Soldos"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4560
   ControlBox      =   0   'False
   Icon            =   "Frm_PostoGrad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   4335
   End
   Begin MSDataGridLib.DataGrid GRID_POST_GRAD 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4335
      _ExtentX        =   7646
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
   Begin VB.CommandButton cmdEditar 
      Caption         =   "E&ditar"
      Height          =   855
      Left            =   1080
      Picture         =   "Frm_PostoGrad.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Editar"
      Top             =   3960
      Width           =   855
   End
   Begin VB.Frame FRA_CADASTRO 
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4335
      Begin VB.TextBox TXT_DESC_POST_GRAD_CATFUNC 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox TXT_POST_GRAD_CATFUNC 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1095
      End
      Begin MSMask.MaskEdBox TXT_SOLDO 
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Soldo:"
         Height          =   195
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição do posto/graduação"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   2220
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Posto/Grad.:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   2040
      Picture         =   "Frm_PostoGrad.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Excluir"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3600
      Picture         =   "Frm_PostoGrad.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Sal&var"
      Height          =   855
      Left            =   120
      Picture         =   "Frm_PostoGrad.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salvar"
      Top             =   3960
      Width           =   855
   End
End
Attribute VB_Name = "frm_PostoGrad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEditar_Click()

cmdSalvar.Enabled = True
cmdEditar.Enabled = False
cmdExcluir.Enabled = True

fra_Cadastro.Enabled = True
TXT_SOLDO.SetFocus

End Sub
Private Sub cmdExcluir_Click()
On Error GoTo Error
If MsgBox("Deseja excluir o registro selecionado?", vbYesNo + vbQuestion, "SisVTR") = vbYes Then
    cmdEditar.Enabled = False
    cmdExcluir.Enabled = False
    cmdSalvar.Enabled = False
    SQL = "select * FROM TAB_GER_POST_GRAD_CATFUNC"
    SQL = SQL + " WHERE POST_GRAD_CATFUNC='" & TXT_POST_GRAD_CATFUNC & "'"
    Set DS = New Recordset
    DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
    DS.Delete
End If
Call LIMPAR
Call Form_Load
MsgBox "Registro deletado.", vbInformation, "SISTRANS"
cmdSalvar.Enabled = True
Exit Sub
Error:
    MsgBox "Erro. O registro não foi excluído,por fazer parte de um relacionamento.", vbOKOnly + vbInformation, "SisTrans"
    cmdSair.SetFocus
End Sub
Private Sub cmdSair_Click()
Unload Me
End Sub
Private Sub cmdSalvar_Click()
If TXT_POST_GRAD_CATFUNC = "" Then
   MsgBox "Digite o Posto/Graduação.", vbInformation, "SISTRANS"
   Exit Sub
End If
If TXT_DESC_POST_GRAD_CATFUNC = "" Then
   MsgBox "Digite a descrição do Posto/Graduação.", vbInformation, "SISTRANS"
   Exit Sub
End If

'If TXT_SOLDO = "" Then
'   MsgBox "Digite o soldo.", vbInformation, "SISTRANS"
'   Exit Sub
'End If

If VarAcao = "Registro Novo" Then
   Call Salvar
   MsgBox "Registro salvo.", vbInformation, "SisVTR"
Else
   Call Salvar
   MsgBox "Registro alterado.", vbInformation, "SisVTR"
End If
Call LIMPAR
Call Form_Load
End Sub
Private Sub LIMPAR()
TXT_POST_GRAD_CATFUNC = ""
TXT_DESC_POST_GRAD_CATFUNC = ""
TXT_SOLDO = ""
End Sub
Private Sub Salvar()
SQL = "select * FROM TAB_GER_POST_GRAD_CATFUNC"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
If VarAcao = "Registro Novo" Then
   DS.AddNew
ElseIf VarAcao = "Registro Velho" Then
   SQL = "select * FROM TAB_GER_POST_GRAD_CATFUNC"
   SQL = SQL + " WHERE POST_GRAD_CATFUNC='" & TXT_POST_GRAD_CATFUNC & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
End If
DS!POST_GRAD_CATFUNC = TXT_POST_GRAD_CATFUNC
DS!DESC_POST_GRAD_CATFUNC = TXT_DESC_POST_GRAD_CATFUNC
DS!SOLDO = TXT_SOLDO
DS.UpdateBatch adAffectAll
End Sub
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
SQL = "select * FROM TAB_GER_POST_GRAD_CATFUNC"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set GRID_POST_GRAD.DataSource = DS
GRID_POST_GRAD.Columns(0).Width = 1500
GRID_POST_GRAD.Columns(1).Width = 3000
GRID_POST_GRAD.Columns(2).Width = 1000
GRID_POST_GRAD.Columns(0).Caption = "Posto/Graduação"
GRID_POST_GRAD.Columns(1).Caption = "Descrição"
GRID_POST_GRAD.Columns(2).Caption = "Soldo"
End Sub
Private Sub GRID_POST_GRAD_Click()
TXT_POST_GRAD_CATFUNC = GRID_POST_GRAD.Columns(0)
TXT_SOLDO = GRID_POST_GRAD.Columns(2)
TXT_DESC_POST_GRAD_CATFUNC = GRID_POST_GRAD.Columns(1)
Call TXT_POST_GRAD_CATFUNC_LostFocus
End Sub
Private Sub TXT_POST_GRAD_CATFUNC_Change()
TXT_DESC_POST_GRAD_CATFUNC = ""
TXT_SOLDO = ""
End Sub
Private Sub TXT_POST_GRAD_CATFUNC_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_DESC_POST_GRAD_CATFUNC_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_POST_GRAD_CATFUNC_Click()
TXT_DESC_POST_GRAD_CATFUNC = ""
TXT_SOLDO = ""
End Sub
Private Sub TXT_POST_GRAD_CATFUNC_LostFocus()
SQL = "select * FROM TAB_GER_POST_GRAD_CATFUNC"
SQL = SQL + " where POST_GRAD_CATFUNC='" & TXT_POST_GRAD_CATFUNC & "'"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
If DS.RecordCount = 1 Then
   
   VarAcao = "Registro Velho"
   
   cmdSalvar.Enabled = False
   cmdEditar.Enabled = True
   cmdExcluir.Enabled = True
   fra_Cadastro.Enabled = False
        
   TXT_POST_GRAD_CATFUNC = DS!POST_GRAD_CATFUNC
   TXT_DESC_POST_GRAD_CATFUNC = DS!DESC_POST_GRAD_CATFUNC
   TXT_SOLDO = DS!SOLDO
Else
    VarAcao = "Registro Novo"
    cmdSalvar.Enabled = True
    cmdEditar.Enabled = False
    cmdExcluir.Enabled = False
    fra_Cadastro.Enabled = True
End If
End Sub

