VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_ListMulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listagem de multa"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   Icon            =   "frm_ListMulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_Sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   7440
      Picture         =   "frm_ListMulta.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Fecha e retorna para tela principal."
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton btn_Editar 
      Caption         =   "E&ditar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   1080
      Picture         =   "frm_ListMulta.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Editar"
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton btn_Salvar 
      Caption         =   "Sal&var"
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      Picture         =   "frm_ListMulta.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salvar"
      Top             =   4560
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid dbg_Listagem 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6376
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
Attribute VB_Name = "frm_ListMulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
