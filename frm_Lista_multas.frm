VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Lista_multas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listagem de multas"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8355
   ControlBox      =   0   'False
   Icon            =   "frm_Lista_multas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btn_Sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   7440
      Picture         =   "frm_Lista_multas.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   4080
      Width           =   855
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
Attribute VB_Name = "frm_Lista_multas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Sair_Click()
Unload Me
End Sub
Private Sub dbg_Listagem_Click()

If DS.RecordCount <> 0 Then
   frm_CadMultaReq.txt_NumTiquete = DS!nr_talao_infr
   frm_CadMultaReq.txt_NumTiquete.Enabled = False
   frm_CadMultaReq.txt_NumTiquete_LostFocus
End If
End Sub
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub
