<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="style.css">
    

<link href="https://fonts.googleapis.com/css2?family=Ubuntu:wght@400;500;700&display=swap" rel="stylesheet">
   
    <title>Exames Veterinarios</title>
</head>
        <%@LANGUAGE = VBScript %><!--#include file="banco5.asp"-->
<%

email = Request.querystring("emailTxt")
xcpf = Request.querystring("xCpf")

if email <> "" then
   x = 1
   do until x = (len(email) + 1)
      letra = mid(email,1,x)
      if letra = "'" then ops = 1
      if letra = "/" then ops = 1
      if letra = "*" then ops = 1
      if letra = ";" then ops = 1
      if letra = "." then ops = 1
      if letra = "," then ops = 1
      if letra = "=" then ops = 1
      if letra = "?" then ops = 1
      if letra = "%" then ops = 1
      if letra = "&" then ops = 1
      if letra = "@" then ops = 1
      if letra = "!" then ops = 1
      if letra = "+" then ops = 1
      if letra = "-" then ops = 1
      if letra = "#" then ops = 1
      if letra = "$" then ops = 1
      x = x + 1
   loop
end if

if xcpf <> "" then
   x = 1
   do until x = (len(xcpf) + 1 )
      letra = mid(xcpf,1,x)
      if letra = "'" then ops = 1
      if letra = "/" then ops = 1
      if letra = "*" then ops = 1
      if letra = ";" then ops = 1
      if letra = "," then ops = 1
      if letra = "=" then ops = 1
      if letra = "?" then ops = 1
      if letra = "%" then ops = 1
      if letra = "&" then ops = 1
      if letra = "@" then ops = 1
      if letra = "!" then ops = 1
      if letra = "+" then ops = 1
      if letra = "-" then ops = 1
      if letra = "#" then ops = 1
      if letra = "$" then ops = 1
      x = x + 1
   loop
end if

if email <> "" and xcpf <> "" then
   best = 1
   xcpf = Cstr(xcfp)
   strq = "Select * from clientes where email = '" & email & "'  and cgc like '" & xcpf & "%'"
   set RstCli = ObjConn.execute(strq)
   cabum = strq

   if RstCli.eof = false then
      session("idCl") = RstCli("CodCli")
      response.redirect "AnimalInterfaceProd.asp?cc=" & RstCli("CodCli")
   else
      wrong = 1
   end if
   RstCli.close
   set RstCli = nothing
end if

response.write "<div class=Container>"
response.write " <div class=section-1><img src=../scr/images/Dogfundo.png ></div>"
response.write "<div class=section-2><form action=CentralProd.asp method=Get>"
response.write "<div class=titleContent><h1>Central de<br>exames</h1> <strong>Preencha os dados de acordo com o infomado no Hospital</strong></div>"
if wrong = 1 then response.write "<h4>Combinação de e-mail ou CPF estão inválidos, tente novamente</h4>"
response.write "<div class=formAlign><div class=inputContainer><label for=emailTxt>Email:</label><input  type=text name=emailTxt value='" & email & "'  placeholder='Seu e-mail'></div>"
response.write "<div class=inputContainer><label for=xCpf>CPF:</label><input name=xCpf size=20 class=style9 type=text onKeyUp=m_CPF(); value='" & xcpf & "'  placeholder='Seu cpf'></div>"
response.write "<button  type=submit name=submit >Entrar <img src=../scr/images/ArrowRight.png></button> </div>"
response.write "</form></div>"
response.write "</div>"
response.write " <div class=section-3><img src=../scr/images/Footer.png ></div> "

ObjConn.close
set RstCliente = nothing
Set ObjConn = nothing
Set SConnString = nothing
Set Strq = nothing
%>
</html>