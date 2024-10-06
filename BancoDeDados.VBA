#Anotações sobre como criar e conectar um banco de dados Access com o VBA

1° ponto - Conectar com o banco de dados e para isso usar a biblioteca Microsoft ActiveX Data Objects 6.1 Library

#Criando a conexão 
Dim gConexao As ADOBD.Connection 'Conector do BancoDeDados

Private Sub IsConectar()
    Dim strConexao As String 
    Set gConexao = New ADODB.Connection
    'Define a conexão 
    strConexao ="Provider=Microsoft.ACE.OLEDB.12.0;DataSource=C:\Users\**\Desktop\Jujuba.accdb;Persist Security Infor=False"

    gConexao.Open strConexao
End Sub 


Provider = Tipo do BD que vai ter a conexão 
DataSource = O local do BD, por ser o access basta colocar o local do arquivo
Persist Security Info = O valor padrão é falso e está relacionado a informações de segurança, como login e senha para conexão com o BD; Tem dois estados: True e False    

Caso o BD tenha uma senha é preico usar a seguinte String de conexão: 

“Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Rieper\Blog Guia do Excel\Excel com banco de dados Access – SQL\SistemaGdE.accdb;Jet OLEDB:Database Password=MyDbPassword;”
Acima trocar o caminho do dataSource pelo caminho do arquivo e MyDbPassWord colocar a sendo do BD


#Desconectar o BancoDeDados 

Private Sub IsDesconectar()
    Is not gConexao Is Nothing then 'Se instanciada o método close é usado e o objeto é setado como Nothing 
        gConexao.Close 
        Set gConexao = Nothing
    End if 
End Sub

*Duas coisas que precisam ser feitas são, conectar (criar a conexão); e Desconectar (Encerrar a conexão) o banco de dados

#Consultando Dados
Para Consultar usar o objeto Recordset 
    Marcar em Ferramentas -> Referências -> Marcar a opção ActiveX Data Objects Recordset 2.8 Library

Public Sub IsListarDados()
    Dim IrS As ADODB.Recordset
    Set Irs = New ADODB.Recordset

    IsConectar 'Verifica se a conexão está ativa, se não estiver conecta.
    Irs.Open "Select * from clientes", gConexao 'Consulta 

    Sheets("Clientes").Range("").ClearContents

    Sheets("Clientes").Cells("1,3").CopyFromRecordset Irs

    If not Irs Is Nothing then 'Se retornar algo é fechado e destruído 
        Irs.Close
        Set Irs = Nothing 
    End if 
    IsDesconectar
End Sub 

**************************************************************************************
#Cadastro 

No banco de dados cada campo foi cadastrado previamente, depois só referenciando eles na consulta SQL.

2 Coisas são necessárias: 1° Execel 2° Access

No Access criar o Banco de dados e configurar ele 
No execel organizar de uma forma onde o usuário possa trabalhar com as informações de forma fácil (inserir, atualizar e remover informaçõe)

Formula do excel para usar formatar uma célula caso ela tenha algo (não vazia) e o valor for impar

#Criando a Conexão 

Private Sub IsConectar() 'Faz a conexão, instancia o objeto de conexão e configura.
    Dim strConexao As String 
    Set gConexao = New ADODB.Connection

    strConexao = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\CaminhoOndeEstáOmeuArquivoDoBancoDeDados;Persist Security Info=False"
    gConexao.Open strConexao
End Sub

É criada uma variável global gConexao que vai ser usada pelo banco de dados

#Desconectando 

Private Sub IsDesconectar() 'Fecha a conexão e destroi o objeto 
    If Not gConexao Is Nothing then
        gConexao.Close
        Set gConexao = Nothing 
    End If
End Sub

#Inserir as informações 

Public Sub IsInserir(ByVal Ides_nome as String, ByVal Ides_logradouro As String, ByVal Inum_nrLogradouro As long, Byval Ides_Bairro As String, ByVal Ides_cidade As String, ByVal Ides_uf As String)
    Dim ISQL As String
    IsConectar
    ISQL = "INSERT INTO Clientes ( des_nome, des_logradouro, num_)" & _
    "VALUES (""" & Ides_nome & """,""" & Ides_logradouro & """, " & Inum_nrLogradouro & ", """ & Ides_bairro & """, """ & Ides_cidade & """, """ & Ides_uf & """)"

    gConexao.Execute ISQL
    IsDesconectar
End Sub

Depois de conectar com o banco de dados as instruções são criadas e executadas de forma direta pela conexão, depois a conexão é encerrada.

#Alterando as Informaçoes 

Public Sub IsAlterar(Byval Ides_nome As String, ByVal Ides_logradouro As String, ByVal Inum_nrLogradouro As Long, ByVal Ides_bairro As String, ByVal Ides_cidade As String, ByVal Ides_uf As String, ByVal IClientes_ID As Long)
    Dim ISQL As String 

    if IClientes_ID > 0 then 
        IsConectar
        ISQL = "UPDATE Clientes SET des_nome = """ & Ides_nome & """, des_logradouro = """ & Ides_logradouro & """, " & "num_logradouro = "  & Inum_nrLogradouro & ", des_bairo= """ & Ides_Bairro & """, des_cidade = """ & Ides_cidade & """, des_uf = """ & Ides_uf & """ where Clientes_ID = " & IClientes_ID 
        gConexao.Execute ISQL
        IsDesconectar
    end if 
End Sub 

Os parâmetros são passados de forma concatenada na variável ISQL e executado diretamente na conexão com o BD
O IClientes_ID é usado como identificação para qual registro será alterado.

#Excluindo os Registros 

Public Sub IsExcluir(ByVal IClientes_ID As Long)
    Dim ISQL

    if IClientes_ID > 0 then
        IsConectar
        ISQL = "DELETE FROM Clientes WHERE Clientes_ID = " & IClientes_ID
        gConexao.Execute ISQL
        IsDesconectar
    end if 
End Sub

#Listar os Dados 
Public Sub IsListarDados()
    Dim Irs As ADODB.Recordset 'Retorna os dados 
    Set Irs = New ADODB.Recordset

    IsConectar 
    Irs.Open "Select * from clientes", gConexao

    Sheets("Clientes").Range("").ClearContents
    Sheets("Clientes").Cells("").CopyFromRecordset Irs 'Os dados são copiados para a planilha, sem precisar de loops 

    If Not Irs Is Nothing Then
        Irs.Close
        Set Irs = Nothing
    End IF
    IsDesconectar
End Sub

#Atualizar Clientes
Public Sub IsAtualizarClientes()
    Dim i As Long
    Dim  IUltimaLinhaAtiva As Long

    'Identifica a última linha preenchida no cadastro
    IUltimaLinhaAtiva = Worksheets("Clientes").Cells(Worksheets("Clientes").Rows.Count, 2).End(xlUp).Row

    For i = 3 to IUltimaLinhaAtiva
        Select Case Cells(i, 8).Value
            Case "Inserir"
                IsInserir Cells(i,2),Cells(i,3),Cells(i,4),Cells(i,5),Cells(i,6),Cells(i,7)
            Case "Alterar"
                IsAlterar Cells(i,2), Cells(i,3), Cells(i,4), Cells(i,5), Cells(i,6), Cells(i, 7), Cells(i, 1)
            Case "Excluir"
                IsExcluir Cells(i, 1) 
            End Select 
    Next i 

    IsListarDados
End Sub
