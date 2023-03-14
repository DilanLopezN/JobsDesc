<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="CentralDeExames.css">
   
    <title>Exames Veterinarios</title>
</head>
<body>
    <main>
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
      response.redirect "AnimalInterfaceTeste.asp?cc=" & RstCli("CodCli")
   else
      wrong = 1
   end if
   RstCli.close
   set RstCli = nothing
end if

response.write "<div class=Layout>"
response.write "<header ><img src=../scr/images/logoPadronizada.png></img></header>"
response.write "<div class=HeaderContainer ><swiper-container class=swiper>"
response.write "<swiper-slide class=adressContainerBaeta><img src=../scr/images/BaetaImg.png><h3>Localização: São Bernardo do Campo <br>Telefone: (11)4336-7185<br>Horário: 24 horas</h3><a href=https://www.drhato.com.br/unidades/baeta-neves/ target=_blank>Saiba mais</a></swiper-slide>"

response.write "<swiper-slide class=adressContainerBaeta><img src=../scr/images/Campestre.png><h3>Localização: Santo André<br>Telefone: (11) 4428-1222 <br>Horário: 24 horas</h3><a href=https://www.drhato.com.br/unidades/campestre/ target=_blank>Saiba mais</a></swiper-slide>"

response.write "<swiper-slide class=adressContainerBaeta><img src=../scr/images/VilaAlto.png><h3>Localização: Santo André - SP<br>Telefone: (11) 4428-1222<br>Horário: Até 22 horas</h3><a href=https://www.drhato.com.br/unidades/vila-alto-santo-andre/ target=_blank>Saiba mais</a></swiper-slide>"

response.write  "</swiper-container></div>"

response.write "<form action=CentralTeste.asp method=Get>"
response.write " <div class=infoContent><strong>Resultado dos exames:  </strong>  <span>Preencha os dados de acordo com o infomado no Hospital</span>"
if wrong = 1 then response.write "<h4>Combinação de e-mail ou CPF estão inválidos, tente novamente</h4>"

response.write "<div class=loginContent><label for=emailTxt>Seu e-mail:</label><div class=inputContainer><img src=../scr/images/EnvelopeSimple.png><input  type=text name=emailTxt value='" & email & "'  placeholder='Seu e-mail'></div>"
response.write "<label for=xCpf>Seu CPF:</label><div class=inputContainer><img src=../scr/images/UserRectangle.png><input name=xCpf size=20 class=style9 type=text onKeyUp=m_CPF(); value='" & xcpf & "'  placeholder='Seu cpf'></div>"
response.write "<tr><td colspan=2 height=3 bgcolor=#f9f9f9 div align=left></td></tr>"
response.write "<button  type=submit name=submit >Entrar</button>"
response.write "</form><tr><td height=18 background=../graphics/BkgGray.jpg colspan=2 div align=right></td></tr>"
response.write "</form><tr><td colspan=2 height=40 bgcolor=#f9f9f9></td></tr>"
response.write "</div>"

ObjConn.close
set RstCliente = nothing
Set ObjConn = nothing
Set SConnString = nothing
Set Strq = nothing
%>
    </main>
   
    <script src="https://cdn.jsdelivr.net/npm/swiper@9/swiper-element-bundle.min.js"></script>
</body>
<!-- 
<footer  id="footerContent" class="footerContent">
        <h3>Copyright - 2023 Dr.Hato  todos os direitos reservados ©</h3>
        <div class="socialLinks">
           <img src="insta.png" alt="instagram">
           <img src="facebook.png" alt="facebook">
           <img src="whatsapp.png" alt="whatsapp">
        </div>
</footer>
-->
</html>