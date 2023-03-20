VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frm_Help 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SisTrans - Ajuda"
   ClientHeight    =   7785
   ClientLeft      =   7050
   ClientTop       =   2205
   ClientWidth     =   6345
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   4200
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":0A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":0CE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":0FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":12AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   1217
      ButtonWidth     =   900
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Voltar"
            Key             =   "Back"
            Object.ToolTipText     =   "Voltar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Próximo"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Índice"
            Key             =   "Home"
            Object.ToolTipText     =   "Inicio - Índice"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Timer tmrBrowser 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   4440
      Top             =   960
   End
   Begin SHDocVwCtl.WebBrowser brwHelp 
      Height          =   7080
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   6360
      ExtentX         =   11218
      ExtentY         =   12488
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Tag             =   "&Address:"
      Top             =   720
      Visible         =   0   'False
      Width           =   5295
   End
End
Attribute VB_Name = "frm_Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()

tmrBrowser.Enabled = True

If (Dir(App.Path & "\help.htm") = "") Then
    MsgBox "O arquivo de ajuda help.htm não foi encontrado. Certifique-se que ele se encontra na pasta base no diretório /help. Caso o problema persista, instale novamente o sistema ou contacte o desenvolvedor.", vbInformation + vbonly, "SisTrans - Ajuda"
    Unload Me
    Exit Sub
End If

brwHelp.Navigate App.Path & "\help.htm"

End Sub


Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
     
On Error Resume Next
     
Select Case Button.Key
    Case "Back"
        brwHelp.GoBack
    Case "Forward"
        brwHelp.GoForward
    Case "Home"
        brwHelp.Navigate App.Path & "\help.htm"
End Select

End Sub

