Attribute VB_Name = "mod_DataRotina"
'Este modulo trata tudo em relação a datas

Public Function fun_DiasUteis(vmod_DataAtual As Date, vmod_QtdDias As Byte) As Date

Dim rs_DataRotina As Recordset
Dim var_DataSomada As Date
Dim var_QtdDias As Byte
Dim var_DiaSemana As Byte

Set rs_DataRotina = New Recordset

var_QtdDias = 1
vmod_QtdDias = vmod_QtdDias + 1

Do While var_QtdDias < vmod_QtdDias

    var_DataSomada = vmod_DataAtual - var_QtdDias
       
    With rs_DataRotina
        .Open "select * from tab_ger_aux_rotina_domingo where data = '" & var_DataSomada & "'", cnConexao, adOpenStatic, adLockOptimistic
        If .RecordCount = 1 Then vmod_QtdDias = vmod_QtdDias + 1
        .Close
    End With
    
    var_DiaSemana = Weekday(var_DataSomada)
    If var_DiaSemana = 7 Or var_DiaSemana = 1 Then vmod_QtdDias = vmod_QtdDias + 1

    var_QtdDias = var_QtdDias + 1
Loop

'Next

fun_DiasUteis = var_DataSomada
        
End Function
