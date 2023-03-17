<!DOCTYPE html>
<html lang="pt-BR">
<head>
   <meta charset="UTF-8">
   <link rel="preconnect" href="https://fonts.googleapis.com">
   <meta http-equiv="X-UA-Compatible" content="IE=edge">
   <meta name="viewport" content="width=device-width, initial-scale=1.0">
   <link href="https://fonts.googleapis.com/css2?family=Ubuntu:wght@400;500;700&display=swap" rel="stylesheet">
   <link rel="stylesheet" href="AnimalInterfaceProd.css">
   <title>Exames Vet</title>
</head>
<main>
   <%
   CodCli = Trim(Request.QueryString("CC"))
msg = Trim(Request.QueryString("msg"))


Set Objconn=Server.CreateObject("ADODB.Connection")
Objconn.mode = 3
sConnString = "DRIVER={Microsoft Access Driver (*.mdb)};" & _ 
"DBQ=" & Server.MapPath("\vet\vet.mdb") & ";"
ObjConn.open (sconnstring), "", "M63b07C42nava"

strq = "Select * from clientes where CodCli=" & CodCli
set RstCadastro = objconn.execute(strq)

strq = "Select * from CadAnimais where CodCli=" & CodCli
set RstAni = objconn.execute(strq)

            if Rstcadastro("CodCli") <> session("idCl") then
               RstCadastro.close
               Set RstCadastro=nothing
               RstAni.close
               Set RstAni = nothing
               set CodCli = nothing
               set carater = nothing
               ObjConn.close
               Set ObjConn = nothing
               response.Redirect("CentralProd.asp")
            end if



response.write "<div class=Container>"
response.write "<div class=section-1><img src=../scr/images/Dogfundo.png></div>"
iF RstAni.eof = false then
   response.write "<div class=section-2>"
   response.write "<h1>Seleção de pets</h1>"
   response.write "<div class=form><p>Selecione seu pet para ver os exames</p>"
   Do Until RstAni.eof = true
   response.write "<div class=petSelect><a href=ExDetailProd.asp?id="& RstAni("CodAnimal") & ">" & RstAni("NomeAnimal") &"</a><img src=../scr/images/ArrowRight.png></div>"
   RstAni.movenext
   loop
   response.write "</div>"
   if Msg<>"" then
   response.write "<tr><td colspan=2><hr><B>"& msg &"</td></tr>"
   end if
   response.write "</div>"
end if
response.write "</div>"
response.write " <div class=section-3><img src=../scr/images/Footer.png ></div> "

RstAni.close
Set RstAni = nothing

RstCadastro.close
Set RstCadastro = nothing

ObjConn.close
set RstCliente = nothing
Set ObjConn = nothing
Set SConnString = nothing
Set Strq = nothing
%>
</main>


</html>






