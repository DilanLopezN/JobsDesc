<!DOCTYPE html>
<html lang="en">
<head>
   <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
   <link rel="preconnect" href="https://fonts.googleapis.com">
   <meta http-equiv="X-UA-Compatible" content="IE=edge">
   <meta name="viewport" content="width=device-width, initial-scale=1.0">
   <title>Seleção de Exames</title>
   <link rel="stylesheet" href="ExDetailProd.css">
   <link href="https://fonts.googleapis.com/css2?family=Ubuntu:wght@400;500;700&display=swap" rel="stylesheet">
</head>
<%

origem = Trim(Request.QueryString("origem"))
id = Trim(Request.QueryString("id"))

Set Objconn=Server.CreateObject("ADODB.Connection")
Objconn.mode = 3
sConnString = "DRIVER={Microsoft Access Driver (*.mdb)};" & _ 
"DBQ=" & Server.MapPath("\vet\vet.mdb") & ";"
ObjConn.open (sconnstring), "", "M63b07C42nava"

strq = "Select * from CadAnimais where CodAnimal=" & id
set RstCadastro = objconn.execute(strq)

If Session("IdCl") <> "" then
   If RstCadastro("CodCli") <> Session("IdCl") then
      response.redirect("Centraldeexames.asp")
   end if
end if

strq = "Select * from ExamesPorConsulta where CodAnimal=" & Id & " and ok=" & -1 & " order by data desc"
set RstExames = objconn.execute(strq)

if RstExames.eof = false then

   session("IdCl") = RstExames("CodCli")


   response.write "<div class=Container>"
   response.write " <div class=section-1><img src=../scr/images/Dogfundo.png ></div>"
   response.write "<div class=section-2>"
   response.write " <a class=backResult href=AnimalInterfaceProd.asp?id="&id&"&origem="&origem&"&CC="&RstCadastro("codcli")&"><img src=../scr/images/ArrowUUpLeft.png><p>Voltar a lista</p></a>"   
   response.write " <h1>Resultado dos exames veterinarios</h1>"
   response.write "<h2>Pet selecionado : "& RstCadastro("NomeAnimal") &"</h2>"
   response.write "<div class=form>"
   iF RstExames.eof = false then
      response.write "<h3>Selecione um exame :</h3>"
      Do Until RstExames.eof = true
         Databr   = RIGHT("0" & day(RstExames("data")),2) & "/" & RIGHT("0" & month(RstExames("data")),2) & "/" & RIGHT("0" & year(RstExames("data")),2)
         response.write "<div class=exameSelect><img src=../scr/images/Frasco.png ><a href=PrintExameProd.asp?CodCli=" & RstCadastro("codcli") & "&text=|"& RstExames("CodExameConsulta") &"| target=_self> "& DataBr & " - " & RstExames("Nomexame") &"</a></div>"
      RstExames.movenext
      loop
   end if
   response.write "</div></div></div>"
   response.write " <div class=section-3><img src=../scr/images/Footer.png ></div>"
   RstExames.close
   Set RstExames = nothing

  
else
   str ="AnimalInterfaceProd.asp?msg=<script>alert('Não existem exames para o pet Selecionado')</script>&CC=" & Session("IdCl")
   rstcadastro.close
   Set RstCadastro = nothing
   ObjConn.close
   Set ObjConn=nothing
   set id = nothing
   Set Origem=nothing
   set strq = nothing
   response.Redirect str
end if

Rstcadastro.close
Set Rstcadastro = nothing

ObjConn.close
set RstCliente = nothing
Set ObjConn = nothing
Set SConnString = nothing
Set Strq = nothing
%>
</html>