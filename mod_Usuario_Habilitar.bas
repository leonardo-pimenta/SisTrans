Attribute VB_Name = "mod_Usuario_Habilitar"
Public vgl_Responsavel As String
Public vgl_Nivel As String

Public Sub pcd_Usuario_Habilitar(vmod_Nivel As String)

If vmod_Nivel = "ADMIN" Then
    'Possui todos os direitos dentro do sistema -Usuário da DN101
    
    frm_Principal.mnu_CGS.Enabled = False
    frm_Principal.tlbMenuEquip.Buttons.Item(12).Enabled = False
    frm_Principal.mnu_OM.Enabled = True
    frm_Principal.mnu_Especialidade.Enabled = True

ElseIf vmod_Nivel = "SUPERVISOR" Then
    
    frm_Principal.tlbMenuEquip.Buttons.Item(12).Enabled = False
    frm_Principal.mnu_CGS.Enabled = False

ElseIf vmod_Nivel = "USER" Then
    
    frm_Principal.mnu_InclirUsuario.Enabled = False
    'frm_Principal.mnu_MultasDetran.Enabled = False
    frm_Principal.tlbMenuEquip.Buttons.Item(12).Enabled = False
    frm_Principal.mnu_CGS.Enabled = False

ElseIf vmod_Nivel = "CGS" Then
    
    frm_Principal.mnu_Cartao.Enabled = False
    frm_Principal.mnu_Multa.Enabled = False
    frm_Principal.mnu_InclirUsuario.Enabled = False
    frm_Principal.mnu_Relatorios.Enabled = False
    frm_Principal.mnu_Manutencao.Enabled = False
    frm_Principal.tlbMenuEquip.Buttons.Item(1).Enabled = False
    frm_Principal.tlbMenuEquip.Buttons.Item(2).Enabled = False
    frm_Principal.tlbMenuEquip.Buttons.Item(4).Enabled = False
    frm_Principal.tlbMenuEquip.Buttons.Item(5).Enabled = False
    frm_Principal.tlbMenuEquip.Buttons.Item(6).Enabled = False
    frm_Principal.tlbMenuEquip.Buttons.Item(7).Enabled = False
    frm_Principal.tlbMenuEquip.Buttons.Item(8).Enabled = False
    frm_Principal.tlbMenuEquip.Buttons.Item(9).Enabled = False
    frm_Principal.tlbMenuEquip.Buttons.Item(10).Enabled = False
    frm_Principal.tlbMenuEquip.Buttons.Item(11).Enabled = False

End If

End Sub
