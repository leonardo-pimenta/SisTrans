Attribute VB_Name = "mod_Conexao"
Public VarAcao As String
Public cnDB As Connection
Public SQL As String
Public DS As Recordset
Public ds1 As Recordset
Public Resp As Integer
Dim rsDB As Recordset
Dim xPosicao As Byte
Public Var_Gl_Cod_Modelo
Public xpath
Public Var_Gl_Cod_Marca
Public vgl_placa As String
Public var_SQL As String
Public var_Resposta As String
Public var_MSG As String
Public vgl_Cor As String
Public vgl_TipoResponsavell As String
Public vgl_TipoResponsavelll As String
Public vgl_TipoResponsavel As String
Public vgl_PostoGrad As String
Public Var_Graduacao As String
Public Var_nome As String
Public vgl_TarjaCartao As String
Public Var_CPF As String
Public Var_CNPJ As String
Public Var_COD As String
Public Var_Gradua��o As String
Public cnConexao As Connection


Public Sub OpenConexao(modSenha, modMDB)

'IMPORTANTE

'N�o esquecer de inserir no formulario que utilizar esta provedure
'um commonDialog com o NAME=cdgConexao

Conectar:

On Error GoTo ErroConectar

'A var Xposicao=1 indica que o erro foi no momento de conexao com o 1� DB
'onde se localiza o path do DB principal do sistema
xPosicao = 1

Open App.Path & "\Conexao.ini" For Input As #1
Input #1, xpath
Close #1

xPosicao = 2

Set cnConexao = New Connection
cnConexao.CursorLocation = adUseClient

'Abre a conexao com o DB principal do sistema j� com o path do arquivo
cnConexao.Open "DRIVER={Microsoft Access Driver (*.mdb)};User ID=; Password=" & modSenha & ";DBQ=" & xpath

Exit Sub

ErroConectar:
        
    Select Case xPosicao
        Case 1
            MsgBox "N�o foi possivel localizar o arquivo Conexao.ini . Certifique-se que este se encontra na pasta base do programa iniciado. Persistindo o erro, contacte o desenvolvedor.", vbInformation + vbOKOnly, "Conex�o"
            End
        Case 2
            MsgBox "O banco de dados do sistema n�o foi localizado. Indique a atual localiza��o a seguir.", vbInformation + vbOKOnly, "Conex�o"
            
            frm_Login.cdg_Conexao.Filter = "" & modMDB & ".mdb|" & modMDB & ".mdb|"
            frm_Login.cdg_Conexao.DialogTitle = "Banco de dados do sistema - " & modMDB & ".mdb"
            frm_Login.cdg_Conexao.ShowOpen
            x = frm_Login.cdg_Conexao.InitDir
                                    
            Open App.Path & "\Conexao.ini" For Output As #1
            Print #1, frm_Login.cdg_Conexao.FileName
            Close #1
            
            GoTo Conectar
    
    End Select
    
    'MsgBox "Para que as altera��es sejam efetuadas, inicie novamente a opera��o", vbInformation + vbOKOnly, "Conex�o"
        
End Sub

