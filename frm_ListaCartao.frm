VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_ListaCartao 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SisTrans - Lista de Cartões Emitidos"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8505
   ControlBox      =   0   'False
   Icon            =   "frm_ListaCartao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btn_Sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   7440
      Picture         =   "frm_ListaCartao.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   5400
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Height          =   135
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   8295
   End
   Begin MSDataGridLib.DataGrid dbg_Listagem 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8705
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
Attribute VB_Name = "frm_ListaCartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Sair_Click()
Unload Me
End Sub
Private Sub dbg_Listagem_DblClick()
If vgl_Nivel = "CGS" Then
   If DS.RecordCount <> 0 Then
      frm_CadMultaReq.txt_NumTiquete = DS!nr_talao_infr
      frm_CadMultaReq.txt_NumTiquete.Enabled = False
      frm_CadMultaReq.txt_NumTiquete_LostFocus
   End If
End If
End Sub
Private Sub dbg_Listagem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
End Sub
Public Sub Form_Load()
Me.Top = 0
Me.Left = 0

SQL = "select * FROM tab_Trans_Cartao order by tarja"

Set DS = New Recordset
DS.Open SQL, cnConexao, adOpenStatic, adLockOptimistic

Set dbg_Listagem.DataSource = DS
'With dbg_Listagem
'     .Columns(0).Visible = False
'     .Columns(4).Visible = False
'     .Columns(6).Visible = False
'     .Columns(9).Visible = False
'     .Columns(10).Visible = False
'     .Columns(11).Visible = False
'     .Columns(12).Visible = False
'     .Columns(13).Visible = False
'     .Columns(14).Visible = False
'     .Columns(15).Visible = False
'     .Columns(1).Width = 1000
'     .Columns(2).Width = 750
'End With

End Sub
