VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_ListaMultaCGS 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SisTrans - Lista de Multas em Espera c/ o Sr CGS"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8370
   Icon            =   "frm_ListaMultaCGS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btn_Sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   7440
      Picture         =   "frm_ListaMultaCGS.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   4320
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   8295
   End
   Begin MSDataGridLib.DataGrid dbg_Listagem 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7011
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
Attribute VB_Name = "frm_ListaMultaCGS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Sair_Click()
Unload Me
End Sub
Private Sub dbg_Listagem_DblClick()
On Error GoTo Error
frm_CadMultaReq.txt_NumTiquete = DS!nr_talao_infr
frm_CadMultaReq.txt_NumTiquete.Enabled = False
frm_CadMultaReq.txt_NumTiquete_LostFocus
Exit Sub
Error:
End Sub
Private Sub dbg_Listagem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
End Sub
Private Sub Form_Activate()
Call frm_Principal.Lista_requerimento_Click
End Sub
Public Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub
