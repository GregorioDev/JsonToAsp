<!--#INCLUDE file="json2.asp" -->

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <title>Exemplo - JSON gerar Objeto ASP</title>        
    </head>
 
    <body>        
        <h1>Exemplo - JSON gerar Objeto ASP</h1>
        <%
            'Declara variavel
            Dim notebook
            Dim variavel_json 
             
            'Variavel com o Json            
            variavel_json = "{""date"":""4/28/2017"",""custType"":""100"",""vehicle"":""1""}"
         
            'Seta Objeto e executa o metodo que converte Json para Objeto ASP       
            Set objeto = JSON.parse(variavel_json)        
         
            response.Write "Nome: "&objeto.date            
            response.Write "<br/>Valor: "&objeto.custType
             
            Set objeto = Nothing     

        %>
    </body>
</html>   