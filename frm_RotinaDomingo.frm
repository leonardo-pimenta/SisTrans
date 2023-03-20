VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_Rotina_Domingo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SisTrans - Rotina de Domingo"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3030
   ControlBox      =   0   'False
   Icon            =   "frm_RotinaDomingo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btn_Excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   960
      Picture         =   "frm_RotinaDomingo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Excluir"
      Top             =   4560
      Width           =   855
   End
   Begin VB.Frame Frame5 
      Caption         =   "Rotinas de Domingo:"
      Height          =   3495
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   2775
      Begin MSDataGridLib.DataGrid dbg_Listagem 
         Height          =   3135
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   4210688
         ForeColor       =   16777215
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
            MarqueeStyle    =   4
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   2775
   End
   Begin VB.CommandButton btn_Sair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   855
      Left            =   2040
      Picture         =   "frm_RotinaDomingo.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton btn_Salvar 
      Caption         =   "S&alvar"
      Default         =   -1  'True
      Height          =   855
      Left            =   120
      Picture         =   "frm_RotinaDomingo.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salvar"
      Top             =   4560
      Width           =   855
   End
   Begin MSMask.MaskEdBox txt_Data 
      Height          =   300
      Left            =   720
      TabIndex        =   8
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   12648447
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   195
      TabIndex        =   4
      Top             =   120
      Width           =   390
   End
End
Attribute VB_Name = "frm_Rotina_Domingo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_TabRotinaDomingo As Recordset
Dim rs_Listagem As Recordset

Private Sub btn_Excluir_Click()

On Error GoTo Error

If MsgBox("Deseja excluir a Rotina de Domingo " & rs_Listagem!Data & " ?", vbYesNo + vbQuestion, "SisTrans") = vbYes Then
    
    rs_TabRotinaDomingo.Open "DELETE * FROM tab_Ger_Aux_Rotina_Domingo WHERE data ='" & rs_Listagem!Data & "'", cnConexao, adOpenStatic, adLockOptimistic
    
    MsgBox "Arquivo excluído.", vbOKOnly + vbInformation, "SisTrans"
    
    Call Form_Activate
    
End If

Exit Sub
Error:
    MsgBox "Erro. O registro não foi excluído.", vbOKOnly + vbInformation, "SisTrans"
    btn_Sair.SetFocus
    
End Sub

Private Sub btn_Sair_Click()
Unload Me
End Sub


Private Sub btn_Salvar_Click()

Dim vCount As Byte

If txt_Data.Text = "" Then
    MsgBox "É necessário entrar com uma Data.", vbOKOnly + vbInformation, "SisTrans"
    txt_Data.SetFocus
    Exit Sub
End If

'Abre a tabela verifica se a senha esta correta, caso sim edita-a
Set rs_TabRotinaDomingo = New Recordset

With rs_TabRotinaDomingo

'''''
    .Open "select * from tab_Ger_Aux_Rotina_Domingo where Data ='" & txt_Data.Text & "'", cnConexao, adOpenStatic, adLockOptimistic

    If .RecordCount = 0 Then
        .AddNew
            !Data = txt_Data.Text
        .UpdateBatch adAffectAll
        MsgBox "Data cadastrado.", vbOKOnly, "SisTrans"
        .Close
        
        Call Form_Activate
    Else
        MsgBox "Esta Data já existe.", vbOKOnly, "SisTrans"
        
        txt_Data.Text = ""
        txt_Data.SetFocus
    End If

End With

End Sub

Private Sub Form_Activate()

Set rs_Listagem = New Recordset
rs_Listagem.Open "select * from tab_Ger_Aux_Rotina_Domingo", cnConexao, adOpenStatic, adLockOptimistic

Set dbg_Listagem.DataSource = rs_Listagem
dbg_Listagem.Columns(0).Width = 1500

txt_Data.Text = "__/__/____"
txt_Data.SetFocus

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub

Private Sub txt_Data_GotFocus()
txt_Data.Text = "__/__/____"
End Sub

Private Sub txt_Data_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

