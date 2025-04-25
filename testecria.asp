<%
' Configurações
Dim remoteUrl, localPath
remoteUrl = "https://i.imgur.com/fgq4vUX.jpeg"  ' URL do arquivo remoto
localPath = Server.MapPath("./arquivos_aleatorios/arquivo.jpeg")  ' Caminho local para salvar

' Cria objeto para fazer a requisição HTTP
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")

On Error Resume Next

' Faz a requisição GET
http.open "GET", remoteUrl, False
http.send

' Verifica se houve erro na requisição
If Err.Number <> 0 Then
    Response.Write "Erro na requisição: " & Err.Description
    Response.End
End If

' Verifica o status da resposta
If http.status = 200 Then
    ' Cria objeto para manipulação de arquivo
    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1  ' 1 = Tipo binário
    
    ' Abre o stream e escreve os dados
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile localPath, 2  ' 2 = Sobrescrever se existir
    stream.Close
    Set stream = Nothing
    
    Response.Write "Arquivo baixado com sucesso!"
Else
    Response.Write "Falha ao baixar. Status: " & http.status & " - " & http.statusText
End If

' Limpa objetos
Set http = Nothing
%>