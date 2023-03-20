VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Marca_Modelo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Marca/Modelo"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fra_Marca 
      Height          =   3495
      Left            =   360
      TabIndex        =   11
      Top             =   480
      Width           =   4935
      Begin MSDataGridLib.DataGrid GRID_MARCA 
         Height          =   2415
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   4260
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
      Begin VB.TextBox TXT_MARCA 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Marca:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdAdicionar 
      Caption         =   "&Adcionar"
      Height          =   855
      Left            =   1440
      Picture         =   "frm_Marca_Modelo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salvar"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "E&ditar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   2280
      Picture         =   "frm_Marca_Modelo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Editar"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Sal&var"
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      Picture         =   "frm_Marca_Modelo.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salvar"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4560
      Picture         =   "frm_Marca_Modelo.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   855
      Left            =   3120
      Picture         =   "frm_Marca_Modelo.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Excluir"
      Top             =   4320
      Width           =   855
   End
   Begin VB.Frame Fra_Modelo 
      Height          =   3495
      Left            =   360
      TabIndex        =   13
      Top             =   480
      Width           =   4935
      Begin VB.ComboBox TXT_TIPO 
         Height          =   315
         ItemData        =   "frm_Marca_Modelo.frx":1412
         Left            =   2640
         List            =   "frm_Marca_Modelo.frx":141C
         TabIndex        =   17
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox TXT_MODELO 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   4695
      End
      Begin MSDataGridLib.DataGrid GRID_MODELO 
         Height          =   1815
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   4695
         _ExtentX        =   8281
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
      Begin MSDataListLib.DataCombo COMBO_MARCA 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "MARCA"
         Text            =   ""
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   2640
         TabIndex        =   16
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Marca:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   495
      End
   End
   Begin MSComctlLib.TabStrip TabSctripMenu 
      Height          =   4095
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   7223
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Marca"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Modelo"
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
Attribute VB_Name = "frm_Marca_Modelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdicionar_Click()

cmdSalvar.Enabled = True
cmdAdicionar.Enabled = False
cmdExcluir.Enabled = False
cmdEditar.Enabled = False
Select Case TabSctripMenu.SelectedItem.Index
    Case 1
         Var_Gl_Cod_Marca = 0
         TXT_MARCA.Enabled = True
         TXT_MARCA.Text = ""
         TXT_MARCA.SetFocus
    Case 2
         Var_Gl_Cod_Modelo = 0
         TXT_MODELO.Enabled = True
         TXT_TIPO.Enabled = True
         COMBO_MARCA.Enabled = True
         TXT_MODELO.Text = ""
         TXT_TIPO.Text = ""
         COMBO_MARCA.Text = ""
         TXT_MODELO.SetFocus
End Select
End Sub
Private Sub cmdEditar_Click()

cmdSalvar.Enabled = True
cmdAdicionar.Enabled = False
cmdExcluir.Enabled = False
cmdEditar.Enabled = False

Select Case TabSctripMenu.SelectedItem.Index
    Case 1
         TXT_MARCA.Enabled = True
    Case 2
         TXT_MODELO.Enabled = True
         COMBO_MARCA.Enabled = True
         TXT_TIPO.Enabled = True
End Select
End Sub
Private Sub cmdExcluir_Click()
On Error GoTo Error
If TXT_MARCA <> "" Then
   SQL = "select MARCA FROM TAB_TRANS_MARCA_VEIC"
   SQL = SQL + " WHERE MARCA='" & TXT_MARCA & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   If MsgBox("Deseja excluir o registro selecionado?", vbYesNo + vbQuestion, "SisVTR") = vbYes Then
      DS.Delete
      DS.UpdateBatch adAffectAll
      Call LIMPAR
      Call Form_Load
      MsgBox "Registro deletado.", vbInformation, "SisVTR"
   End If
End If

If COMBO_MARCA.Text <> "" And TXT_MODELO.Text <> "" Then
   SQL = "select * FROM TAB_TRANS_MODELO_VEIC,"
   SQL = SQL + " TAB_TRANS_MARCA_VEIC"
   SQL = SQL + " where MODELO='" & TXT_MODELO & "'"
   SQL = SQL + " AND TAB_TRANS_MODELO_VEIC.COD_MARCA=TAB_TRANS_MARCA_VEIC.COD"
   SQL = SQL + " AND TAB_TRANS_MARCA_VEIC.MARCA='" & COMBO_MARCA & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      varmodelo = DS(0)
      
      SQL = "select * FROM TAB_TRANS_MODELO_VEIC"
      SQL = SQL + " where cod=" & varmodelo
      Set DS = New Recordset
      DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
      If MsgBox("Deseja excluir o registro selecionado?", vbYesNo + vbQuestion, "SisVTR") = vbYes Then
      
         DS.Delete
         DS.UpdateBatch adAffectAll
         Call LIMPAR
         Call Form_Load
         MsgBox "Registro deletado.", vbInformation, "SisVTR"
      End If
End If
Exit Sub
Error:
  MsgBox "Não é possível excluir o registro, ele pode fazer parte de um ou mais relacionamentos.Ex.: multa x veículo, veículo x proprietário. Caso seja realmente necessária a exclusão, contate o administrador do Banco de Dados.", vbOKOnly + vbInformation, "SisTrans"
  cmdSair.SetFocus
End Sub
Private Sub cmdSair_Click()
Unload Me
End Sub
Private Sub cmdSalvar_Click()

If TXT_MARCA <> "" Then

   SQL = "select * FROM TAB_TRANS_MARCA_VEIC"
   SQL = SQL + " WHERE COD=" & Var_Gl_Cod_Marca
   SQL = SQL + " OR  MARCA='" & TXT_MARCA & "'"
   
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   If DS.RecordCount = 1 Then
      DS!MARCA = TXT_MARCA
   Else
      DS.AddNew
      DS!MARCA = TXT_MARCA
   End If
   
   DS.UpdateBatch adAffectAll
End If

If TXT_MODELO <> "" Then
   If COMBO_MARCA = "" Then
      MsgBox "Selecine a marca.", vbInformation, "SISTRANS"
      Exit Sub
   End If
   SQL = "select * FROM TAB_TRANS_MARCA_VEIC"
   SQL = SQL + " WHERE MARCA='" & COMBO_MARCA.Text & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   VARCODMARCA = DS!cod
      
   SQL = "select * FROM TAB_TRANS_MODELO_VEIC"
   SQL = SQL + " WHERE COD=" & Var_Gl_Cod_Modelo
   SQL = SQL + " OR  MODELO='" & TXT_MODELO & "'"
   Set DS = New Recordset
   DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
   If DS.RecordCount = 1 Then
      DS!modelo = TXT_MODELO
      DS!COD_MARCA = VARCODMARCA
      DS!TIPO = TXT_TIPO
   Else
      DS.AddNew
      DS!modelo = TXT_MODELO
      DS!COD_MARCA = VARCODMARCA
      DS!TIPO = TXT_TIPO
   End If
   DS.UpdateBatch adAffectAll
End If
Call LIMPAR
Call Form_Load
End Sub
Private Sub LIMPAR()
TXT_MARCA = ""
TXT_MODELO = ""
TXT_TIPO = ""
COMBO_MARCA.Text = ""
SQL = "select MARCA FROM TAB_TRANS_MARCA_VEIC"
SQL = SQL + " order by MARCA"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set COMBO_MARCA.RowSource = DS
End Sub
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
'GRID_MARCA
SQL = "select MARCA,COD FROM TAB_TRANS_MARCA_VEIC"
SQL = SQL + " order by MARCA"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set GRID_MARCA.DataSource = DS
GRID_MARCA.Columns(0).Width = 4050
GRID_MARCA.Columns(1).Visible = False

'GRID_MODELO
SQL = "select MODELO,MARCA,TIPO,TAB_TRANS_MARCA_VEIC.COD,TAB_TRANS_MODELO_VEIC.COD"
SQL = SQL + " FROM TAB_TRANS_MARCA_VEIC,TAB_TRANS_MODELO_VEIC"
SQL = SQL + " WHERE TAB_TRANS_MARCA_VEIC.COD=TAB_TRANS_MODELO_VEIC.COD_MARCA "
Set ds1 = New Recordset
ds1.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set GRID_MODELO.DataSource = ds1
GRID_MODELO.Columns(0).Width = 1500
GRID_MODELO.Columns(1).Width = 1500
GRID_MODELO.Columns(2).Width = 1500
GRID_MODELO.Columns(3).Visible = False
GRID_MODELO.Columns(4).Visible = False

'Combo Marca
SQL = "select MARCA FROM TAB_TRANS_MARCA_VEIC"
SQL = SQL + " order by MARCA"
Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic
Set COMBO_MARCA.RowSource = DS

TXT_MARCA.Enabled = False
TXT_MODELO.Enabled = False
TXT_TIPO.Enabled = False
COMBO_MARCA.Enabled = False

cmdSalvar.Enabled = False
cmdAdicionar.Enabled = True
cmdEditar.Enabled = False
cmdExcluir.Enabled = False


End Sub

Private Sub GRID_MARCA_Change()
TXT_MARCA = GRID_MARCA.Columns(0)
Var_Gl_Cod_Marca = GRID_MARCA.Columns(1)

cmdExcluir.Enabled = True
cmdEditar.Enabled = True
cmdAdicionar.Enabled = False
cmdSalvar.Enabled = False

End Sub

Private Sub GRID_MARCA_DBLClick()

TXT_MARCA = GRID_MARCA.Columns(0)
Var_Gl_Cod_Marca = GRID_MARCA.Columns(1)

cmdExcluir.Enabled = True
cmdEditar.Enabled = True
cmdAdicionar.Enabled = False
cmdSalvar.Enabled = False

End Sub

Private Sub GRID_MODELO_Change()
TXT_MODELO = GRID_MODELO.Columns(0)
TXT_TIPO = GRID_MODELO.Columns(2)
COMBO_MARCA.Text = GRID_MODELO.Columns(1)
Var_Gl_Cod_Modelo = GRID_MODELO.Columns(4)

cmdExcluir.Enabled = True
cmdEditar.Enabled = True
cmdAdicionar.Enabled = False
cmdSalvar.Enabled = False
End Sub
Private Sub GRID_MODELO_DBLCLICK()
TXT_MODELO = GRID_MODELO.Columns(0)
TXT_TIPO = GRID_MODELO.Columns(2)
COMBO_MARCA.Text = GRID_MODELO.Columns(1)
Var_Gl_Cod_Modelo = GRID_MODELO.Columns(4)

cmdExcluir.Enabled = True
cmdEditar.Enabled = True
cmdAdicionar.Enabled = False
cmdSalvar.Enabled = False

End Sub
Private Sub TabSctripMenu_Click()
Select Case TabSctripMenu.SelectedItem.Index
    Case 1
          Fra_Marca.Visible = True
          Fra_Modelo.Visible = False
    Case 2
          Fra_Modelo.Visible = True
          Fra_Marca.Visible = False
End Select
End Sub
Private Sub TXT_MARCA_Change()
If TXT_MARCA = "" Then
   Call Form_Load
End If
End Sub
Private Sub TXT_MARCA_Click()
If TXT_MARCA = "" Then
   Call Form_Load
End If
End Sub
Private Sub TXT_MARCA_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_MODELO_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TXT_MODELO_Change()
If TXT_MODELO.Text = "" Then
   Call Form_Load
End If
End Sub
Private Sub TXT_MODELO_Click()
If TXT_MODELO.Text = "" Then
   Call Form_Load
End If
End Sub

Private Sub TXT_TIPO_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
