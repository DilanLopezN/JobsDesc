<script type="text/javascript" src="verificaCpfexames.js" ></script>
<style TYPE ="text/css">
A{text-decoration:none;font:Verdana;}
p{font-family:Arial;color:#000000;font-size:100%;}
.style9{font-size:9px;font-family:Verdana;color=#000000;}
</style><html><head><title>Sistema Veterinário - Exames Veterinários</title><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
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
      response.redirect "AnimalInterface.asp?cc=" & RstCli("CodCli")
   else
      wrong = 1
   end if
   RstCli.close
   set RstCli = nothing
end if

response.write "<table border=0 width=800 class=style9>"
response.write "<tr><td width=800 colspan=2 div align=center><img src=../scr/images/LogoNovo.jpg border=0></img></td></tr>"
response.write "<tr><td width=800 colspan=2><font color=#066e12>Baeta Neves - (11) 4336-7185 São Bernardo do Campo - SP - Rua Thales dos Santos Freire,136 Vila Baeta Neves. 24 horas</td></tr>"
response.write "<tr><td width=800 colspan=2><font color=#066e12>Campestre - (11) 4428-1222 Santo André - SP - Av. Dom Pedro II,3.309 Bairro Campestre. 24 horas</td></tr>"
response.write "<tr><td width=800 colspan=2><font color=#066e12>Vila Alto de Sto André - (11) 4200-1160 Santo André - SP - Av. Martim Francisco,802 Vila Alto de Santo André.</td></tr>"

response.write "<form action=centraldeexames.asp method=Get>"
response.write "<tr><td colspan=2 div align=center><br><hr><B>Resultados de exames veterinários.</b><div align=left><br><b>Preencha os dados de acordo com o infomado no Hospital</b><hr></td></tr>"
if wrong = 1 then response.write "<tr><td colspan=2 div align=center><br><font color=#FF0000>Combinação de e-mail / CPF inválidos</td></tr>"

response.write "<tr><td bgcolor=#f9f9f9 div align=left>E-mail</td><td><input class=style9 size=50 type=text name=emailTxt value='" & email & "'></input></td></tr></td></tr>"
response.write "<tr><td bgcolor=#f9f9f9 div align=left>CPF</td><td><input name=xCpf size=20 class=style9 type=text onKeyUp=m_CPF(); value='" & xcpf & "'></input></td></tr>"
response.write "<tr><td colspan=2 height=3 bgcolor=#f9f9f9 div align=left></td></tr>"
response.write "<tr><td colspan=2 div align=right><input class=style9 type=submit name=submit value=Entrar></input><br><hr></td></tr>"
response.write "</form><tr><td height=18 background=../graphics/BkgGray.jpg colspan=2 div align=right></td></tr>"
response.write "</form><tr><td colspan=2 height=40 bgcolor=#f9f9f9></td></tr>"
response.write "</table>"

ObjConn.close
set RstCliente = nothing
Set ObjConn = nothing
Set SConnString = nothing
Set Strq = nothing
%>