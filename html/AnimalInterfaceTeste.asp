<!DOCTYPE html>
<html lang="en">
<head>
   <meta charset="UTF-8">
   <meta http-equiv="X-UA-Compatible" content="IE=edge">
   <meta name="viewport" content="width=device-width, initial-scale=1.0">
   <link rel="stylesheet" href="AnimalInterface.css">
   <title>Exames Vet</title>
</head>
<body>
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
               response.Redirect("CentralTeste.asp")
            end if

response.write "<header>"
response.write "<img src=../scr/images/logoPadronizada.png> </header>"
response.write "<div class=mainContent><table width=804 class=style9 border=0 bgcolor=#f7f7f7>"
response.write "<h2>Resultados de exames veterinários.</h2>"
if Msg<>"" then
   response.write "<tr><td colspan=2><hr><B>"& msg &"</td></tr>"
end if
iF RstAni.eof = false then
   response.write "<div class=selectMenu><h4 >Selecione seu pet:</h4> <img src=../scr/images/ArrowSquareDown.png ></div>"
   Do Until RstAni.eof = true
      if bgcolor="f7f7f7" then bgcolor="ffffff" else bgcolor="f7f7f7"
      response.write "<div class=selectPet><img src=../scr/images/HeartRed.png > <a href=ExDetailTeste.asp?id="& RstAni("CodAnimal") & ">" & RstAni("NomeAnimal") &"</a></div>"
'<img src=../scr/images/lilstar.png border=0></img>
   RstAni.movenext
   loop
end if
response.write "</table></div>"

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
</body>
</html>






