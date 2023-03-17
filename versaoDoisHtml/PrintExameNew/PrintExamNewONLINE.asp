<style TYPE ="text/css">
A{text-decoration: none;font:Verdana;p{font-family: Arial, sans-serif; color: #000000; font-size: 100%%;}}
.style9{font-size:9px;font-family:Verdana;color=#000000;}
#Layer1{display:block;width:800;height:18;position:absolute;left:0;top:0;z-index=1;}
#Layer2{display:block;width:800;height:18;position:absolute;left:0;top:120;z-index=1;}

</style><html><head><title>Hospital Veterinário Dr Hato - Exame</title><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%@LANGUAGE = VBScript %><!--#include file="banco5.asp"--><%
text = Trim(Request.QueryString("text"))
origem = Trim(Request.QueryString("origem"))
CodCli=Trim(Request.QueryString("CodCli"))

if len(text) > 0 then
   x = 1
   y = len(text)
   Caracter = ""
   ok = 0
   Do until x > ( y )
      caracter = mid((text),x,1)
      If caracter = "|" then
      else
         CodEx = CodEx & caracter
      end if
      If caracter = "|" then
         If CodEx <> "" then
            Response.Expires = 0
            Id = CodEx
            strQ = "select * from ConsultaExamePorConsulta where codexameconsulta=" & id
            set ObjRS=ObjConn.execute(strQ)

            if ObjRs("CodCli") <> session("idCl") then
               ObjRs.close
               Set ObjRs = nothing
               text  = ""
               set y = nothing
               set x = nothing
               set CodEx=nothing
               set carater = nothing
               ObjConn.close
               Set ObjConn = nothing
               response.Redirect("centraldeexames.asp")
            end if

            response.write "<div id=Layer1>"
            response.write "<table width=800 border=0 nowrap class=style9><tr><td width=80>"
            response.write "<tr><td><img src=../scr/images/LogoNovo.jpg border=0></img></td></tr>"
            response.write "<tr><td width=800 colspan=2><font color=#066e12>Baeta Neves - (11) 4336-7185 São Bernardo do Campo - SP - Rua Thales dos Santos Freire,136 Vila Baeta Neves. 24 horas</td></tr>"
            response.write "<tr><td width=800 colspan=2><font color=#066e12>Campestre - (11) 4428-1222 Santo André - SP - Av. Dom Pedro II,3.309 Bairro Campestre. 24 horas</td></tr>"  
            response.write "<tr><td width=800 colspan=2><font color=#066e12>Vila Alto de Sto André - (11) 4200-1160 Santo André - SP - Av. Martim Francisco,802 Vila Alto de Santo André.</td></tr>"
            response.write "<tr><td colspan=2><a href=ExDetail.asp?id="& ObjRs("CodAnimal") & "><b>Voltar à lista de exames.</a></td></tr></table></div>"
            response.write "<div id=Layer2>"
            response.write "<table width=800 border=0 nowrap class=style9><tr><td width=80>"
            CodigoChave = ObjRs("codexameconsulta")
            codexame = objrs("codexame")
            strq = "select * from exames where codexame=" & ObjRs("CodExame")
            set rstExames = ObjConn.execute(strq)
            strq = "select * from Clientes where codcli= " & ObjRs("CodCli")
            set RstCadastro = ObjConn.execute(strq)
            strq = "select * from CadAnimais where CodAnimal= " & ObjRs("CodAnimal")
            set ObjRsAnimais = ObjConn.execute(strq)
            strq = "Select * from Users where UserCode=" & ObjRs("veterinário")
            set RstUsers = ObjConn.execute(strq)

            nomenaimpressao = rstExames("nomenaimpressao")
            Objrs.movefirst
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Response.write "<tr><td height=1 bgcolor=#cccccc colspan=2></td></tr>"
            Response.write "<tr><td bgcolor=#f9f9f9 div align=center Colspan=2><Font Face=Verdana size=1>Resultado de Exame</td></tr><tr><td>"
            Response.write "<Font Face=Verdana size=1>Cliente</td><td width=600><Font Face=Verdana size=1> " & RstCadastro("nome") & "</b></font></td></tr>"
            Response.write "<td width=90><Font Face=Verdana size=1>Animal</td><td><Font Face=Verdana size = 1>" & ObjrsAnimais("nomeanimal") & ", " & ObjrsAnimais("Tipo") & ", " & ObjrsAnimais("raça") & ", "


                  if isdate(ObjRsAnimais("nascimento"))=true then
                   dias = DateDiff("d",ObjRsAnimais("nascimento"),ObjRs("data"))
                   anos = Int(dias / 365)
                   resto = int(12 * ((dias / 365) - Int(dias / 365)))
                   If anos > 1 then
                      if resto > 1 then Idade = anos & " anos e " & resto & " meses "
                      if resto = 1 then Idade = anos & " anos e " & resto & " mês "
                      if resto < 1 then Idade = anos & " anos."
                   end if
                   If anos = 1 then
                      if resto > 1 then Idade = anos & " ano e " & resto & " meses "
                      if resto = 1 then Idade = anos & " ano e " & resto & " mês "
                      if resto < 1 then Idade = anos & " ano."
                   end if
                   if anos  = 0 then
                      if resto > 1 then Idade = resto & " meses "
                      if resto = 1 then Idade = resto & " mês "
                      IF RESTO = 0 then iDADE = "Menos de um Mês"
                   end if
                else
                   Idade="Erro"
                end if

            Response.write Idade & ", " & ObjrsAnimais("Sexo") & ", Código: " & ObjrsAnimais("codanimal") &" </td></tr>"
            brData = RIGHT("0" & day(Objrs("data")),2) & "/" & RIGHT("0" & Month(Objrs("data")),2) & "/" & RIGHT("0" & Year(Objrs("data")),2)
            Response.write "<tr><td><Font Face=Verdana size = 1>Data</td><td><Font Face=Verdana size = 1> " & BrData & "</b></font></td></tr>"
            Response.write "<td><Font Face=Verdana size= 1>Exame</td><td><Font Face=Verdana size = 1> " & RstExames("nomeexame") & "</b></font></td></tr>"
            Response.write "<tr><td ><Font Face=Verdana size=1>Solicitante</td><td><Font Face=Verdana size = 1> " & RstUsers("NomeCOmpleto") & " / CRMV : " & RstUsers("Crm") & "</b></font></td></tr>"
            response.write "<tr><td width=700 height=1 bgcolor=#eae6e6 colspan=2></td></tr>"
            Response.write "</table>"
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''


            If RstExames("TextoLaudo") = 0 then
               strq = "select * from PartesExames where codexame= " & codexame '& " order by coddetalhe"
               set rstParts = objconn.execute(strq)
               rstparts.movefirst
               strq ="select * from Detalhes_Consulta_Exames where codparte <>" & rstParts("codparte") & " and autocode_exame = " & codexame & " order by caracteristica_code"
               set rstDetParts2 = objconn.execute(strq)
               rstParts.movefirst
               response.write "<table width=800 border=0 class=style9>"
               response.write "<tr>"
               response.write "<td width=150><b>" & rstParts("NomeParte") & "</b></td>"
               response.write "<td colspan=2 width=160 div align=center>Resultados</td>"
               response.write "<td colspan=2 width=160 div align=center>Unidades</td>"
               strq = "select * from Detalhes_Consulta_Exames where codparte=" & rstParts("codparte") & " and autocode_exame = " & codigoChave & " order by caracteristica_code"
               set rstDetParts = objconn.execute(Strq)
               strq = "Select * from detalhespartes where coddetalhe=" & rstDetParts("caracteristica_code")
               set rstOrgnlData = objconn.execute(strq)
               testL1 = mid(rstOrgnlData("Nome1"),1,1)
               If testL1 <> "#" and testL1 <> "@" then
                  If int(dias) <= int(rstOrgnlData("Logica1")) then sho1 = 1
                  If int(dias) > int(rstOrgnlData("Logica1")) and int(dias) < int(rstOrgnlData("Logica2")) then sho2 = 1
                  If int(dias) > int(rstOrgnlData("Logica2")) and int(dias) < int(rstOrgnlData("Logica3")) then sho3 = 1
               end if
               If testL1 = "@" then
                  strq = "Select * from SelecionaCabecalho where CodParte=" & RstParts("CodParte")
                  set rstSelCab = ObjConn.execute(strq)
                  do until rstselcab.eof = true
                     If RstSelCab("Coluna1")  = ObjRsAnimais("Tipo") then sho1 = 1
                     If RstSelCab("Coluna2")  = ObjRsAnimais("Tipo") then sho2 = 1
                     If RstSelCab("Coluna3")  = ObjRsAnimais("Tipo") then sho3 = 1
                     rstSelCab.movenext
                  loop
               end if
               If TestL1 = "#" then
                  If rstOrgnlData("Nome1") <> "" then sho1 = 1
                  If rstOrgnlData("Nome2") <> "" then sho2 = 1
                  If rstOrgnlData("Nome3") <> "" then sho3 = 1
               end if
               if sho1 = 1 then
                  if testL1 = "#" then response.write "<td nowrap width=160 colspan=2 div align=center>" & replace(rstOrgnlData("nome1"),"#","") & "</td>"
                  if testL1 = "@" then response.write "<td nowrap width=160 colspan=2 div align=center>" & replace(rstOrgnlData("nome1"),"@","") & "</td>"
                  if testL1 <> "@" and TestL1 <> "#" then response.write "<td nowrap width=160 colspan = 2 div align=center>" & rstOrgnlData("nome1") & "</td>"
               end if
               if sho2 = 1 then
                  if testL1 = "#" then response.write "<td nowrap width=160 colspan=2 div align=center>" & replace(rstOrgnlData("nome2"),"#","") & "</td>"
                  if testL1 = "@" then response.write "<td nowrap width=160 colspan=2 div align=center>" & replace(rstOrgnlData("nome2"),"@","") & "</td>"
                  if testL1 <> "@" and TestL1 <> "#" then response.write "<td nowrap width=160 colspan=2 div align=center>" & rstOrgnlData("nome2") & "</td>"
               end if
               if sho3 = 1 then
                  if testL1 = "#" then response.write "<td nowrap width=160 colspan=2 div align=center>" & replace(rstOrgnlData("nome3"),"#","") & "</td>"
                  if testL1 = "@" then response.write "<td nowrap width=160 colspan=2 div align=center>" & replace(rstOrgnlData("nome3"),"@","") & "</td>"
                  if testL1 <> "@" and TestL1 <> "#" then response.write "<td nowrap width=160 colspan=2 div align=center>" & rstOrgnlData("nome3") & "</td>"
               end if
               response.write "</tr><tr>"
               response.write "<td width=150>Característica</td>"
               response.write "<td div align=center width=80>Absoluto</td>"
               response.write "<td div align=center width=80>Relativo</td>"
               response.write "<td div align=center width=80>Un. Abs.</td>"
               response.write "<td div align=center width=80>Un. Rel.</td>"
               response.write "<td div align=center width=80>Absoluto</td>"
               response.write "<td div align=center width=80>Relativo</td>"
               response.write "</tr>"
               do until rstDetParts.eof = true
                  response.write "<tr>"
                  strq ="Select * from detalhespartes where coddetalhe="& rstDetParts("caracteristica_code")
                  set rstOrgnlData = objconn.execute(strq)
                  response.write "<td nowrap width=150>" & rstOrgnlData("caracteristica") &"</td>"
                  response.write "<td div align=center nowrap width=80>" & rstDetParts("resultadoAbsoluto") & "</td>"
                  response.write "<td div align=center nowrap width=80>" & rstDetParts("resultadoRelativo") & "</td>"
                  response.write "<td div align=center nowrap width=80>" & rstOrgnlData ("unidadeAbsoluto")&"</td>"
                  response.write "<td div align=center nowrap width=80>" & rstOrgnlData ("unidadeRelativo")&"</td>"
                  if sho1 = 1 then
                  response.write "<td div align=center nowrap=nowrap width=100>" & rstOrgnlData ("minimo")& "</td>"
                  response.write "<td div align=center nowrap=nowrap width=100>" & rstOrgnlData ("maximo")& "</td>"
                  end if
                  if sho2 = 1 then
                  response.write "<td div align=center nowrap=nowrap width=100>" & rstOrgnlData ("minimo2")&"</td>"
                  response.write "<td div align=center nowrap=nowrap width=100>" & rstOrgnlData ("maximo2")&"</td></tr>"
                  end if
                  if sho3 = 1 then
                  response.write "<td div align=center nowrap=nowrap width=100>" & rstOrgnlData ("minimo3")&"</td>"
                  response.write "<td div align=center nowrap=nowrap width=100>" & rstOrgnlData ("maximo3")&"</td></tr>"
                  end if
                  dias = DateDiff("d",RstCadastro("nascimento"),date)
                  rstDetParts.movenext
               loop
               response.write "</table>"
               response.write "<table width=800 class=style9>"
               response.write "<tr><td HEIGHT=1 BGCOLOR=#cccccc width=700></td></tr>"
               response.write "<tr><td><b>Observações : </b>" & ObjRs("Obs1") & "</td></tr>"
               response.write "<tr><td HEIGHT=1 BGCOLOR=#cccccc width=700></td></tr>"
               response.write "</Table>"
               rstParts.movenext
               if rstparts.eof = false then
                  response.write "<table nowrap border=0 width=800 class=style9>"
                  response.write "<tr>"
                  response.write "<td width=150><b>" & rstParts("NomeParte") & "</td>"
                  response.write "<td width=160 colspan=2 div align=center>Resultados</td>"
                  response.write "<td td width=160 colspan=2 div align=center>Unidades</td>"
                  strq = "select * from Detalhes_Consulta_Exames where codparte=" & rstParts("codparte") & " and autocode_exame = " & codigoChave & " order by caracteristica_code"
                  set rstDetParts2 = objconn.execute(Strq)
                  strq = "Select * from detalhespartes where coddetalhe=" & rstDetParts2("caracteristica_code")
                  set rstOrgnlData = objconn.execute(strq)
                  testL1 = mid(rstOrgnlData("Nome1"),1,1)
                  sho1 = 0
                  sho2 = 0
                  sho3 = 0
                  If testL1 <> "#" and testL1 <> "@" then
                     If int(dias) <= int(rstOrgnlData("Logica1")) then sho1 = 1
                     If int(dias) > int(rstOrgnlData("Logica1")) and int(dias) < int(rstOrgnlData("Logica2")) then sho2 = 1
                     If int(dias) > int(rstOrgnlData("Logica2")) and int(dias) < int(rstOrgnlData("Logica3")) then sho3 = 1
                  end if
                  If testL1 = "@" then
                     strq = "Select * from SelecionaCabecalho where CodParte=" & RstParts("CodParte")
                     set rstSelCab = ObjConn.execute(strq)
                     do until rstSelCab.eof = true
                        If RstSelCab("Coluna1")  = ObjRsAnimais("Tipo") then sho1 = 1
                        If RstSelCab("Coluna2")  = ObjRsAnimais("Tipo") then sho2 = 1
                        If RstSelCab("Coluna3")  = ObjRsAnimais("Tipo") then sho3 = 1
                        rstSelCab.movenext
                     loop
                  end if
                  If TestL1 = "#" then
                     If rstOrgnlData("Nome1") <> "" then sho1 = 1
                     If rstOrgnlData("Nome2") <> "" then sho2 = 1
                     If rstOrgnlData("Nome3") <> "" then sho3 = 1
                  end if
                  if sho1 = 1 then
                     if testL1 = "#" then response.write "<td nowrap width=160 colspan=2 div align=center>" & replace(rstOrgnlData("nome1"),"#","") & "</td>"
                     if testL1 = "@" then response.write "<td nowrap width=160 colspan=2 div align=center>" & replace(rstOrgnlData("nome1"),"@","") & "</td>"
                     if testL1 <> "@" and TestL1 <> "#" then response.write "<td nowrap width=160 colspan=2 div align=center>" & rstOrgnlData("nome1") & "</td>"
                  end if
                  if sho2 = 1 then
                     if testL1 = "#" then response.write "<td nowrap width=160 colspan=2 div align=center>" & replace(rstOrgnlData("nome2"),"#","") & "</td>"
                     if testL1 = "@" then response.write "<td nowrap width=160 colspan=2 div align=center>" & replace(rstOrgnlData("nome2"),"@","") & "</td>"
                     if testL1 <> "@" and TestL1 <> "#" then response.write "<td nowrap width=160 colspan=2 div align=center>" & rstOrgnlData("nome2") & "</td>"
                  end if
                  if sho3 = 1 then
                     if testL1 = "#" then response.write "<td nowrap width=160 colspan=2 div align=center>" & replace(rstOrgnlData("nome3"),"#","") & "</td>"
                     if testL1 = "@" then response.write "<td nowrap width=160 colspan=2 div align=center>" & replace(rstOrgnlData("nome3"),"@","") & "</td>"
                     if testL1 <> "@" and TestL1 <> "#" then response.write "<td nowrap width=160 colspan=2 div align=center>" & rstOrgnlData("nome3") & "</td>"
                  end if
                  response.write "</tr><tr>"
                  response.write "<td width=150>Característica</td>"
                  response.write "<td div align=center width=80>Absoluto</td>"
                  response.write "<td div align=center width=80>Relativo</td>"
                  response.write "<td div align=center width=80>Un. Abs.</td>"
                  response.write "<td div align=center width=80>Un. Rel.</td>"
                  response.write "<td div align=center width=80>Absoluto</td>"
                  response.write "<td div align=center width=80>Relativo</td>"
                  response.write "</tr>"
                  do until rstDetParts2.eof = true
                     response.write "<tr>"
                     strq ="Select * from detalhespartes where coddetalhe= " & rstDetParts2("caracteristica_code")
                     set rstOrgnlData = objconn.execute(strq)
                     response.write "<td nowrap width=150>" & rstOrgnlData("caracteristica") &"</td>"
                     response.write "<td div align=center nowrap width=80>" & rstDetParts2("resultadoAbsoluto") & "</td>"
                     response.write "<td div align=center nowrap width=80>" & rstDetParts2("resultadoRelativo") & "</td>"
                     response.write "<td div align=center nowrap width=80>" & rstOrgnlData ("unidadeAbsoluto")&"</td>"
                     response.write "<td div align=center nowrap width=80>" & rstOrgnlData ("unidadeRelativo")&"</td>"
                     if sho1 = 1 then
                        response.write "<td div align=center nowrap=nowrap width=100>" & rstOrgnlData ("minimo")& "</td>"
                        response.write "<td div align=center nowrap=nowrap width=100>" & rstOrgnlData ("maximo")& "</td>"
                     end if
                     if sho2 = 1 then
                        response.write "<td div align=center nowrap=nowrap width=100>" & rstOrgnlData ("minimo2")&"</td>"
                        response.write "<td div align=center nowrap=nowrap width=100>" & rstOrgnlData ("maximo2")&"</td></tr>"
                     end if
                     if sho3 = 1 then
                        response.write "<td div align=center nowrap=nowrap width=100>" & rstOrgnlData ("minimo3")&"</td>"
                        response.write "<td div align=center nowrap=nowrap width=100>" & rstOrgnlData ("maximo3")&"</td></tr>"
                     end if
                     rstdetparts2.movenext
                  loop
                  response.write "<tr><td colspan=20 HEIGHT=1 BGCOLOR=CCCCCC width=700>"
                  response.write "</table>"
'                  strq = "Select * from obs2 where codexameconsulta=" & objrs("codexameconsulta")
'                  set objRsObs2 = objconn.execute(strq)
                  response.write "<table width=800 class=style9>"
                  response.write "<tr><td><b>Observações : </b>" & ObjRs("Obs2") & "</td></tr>"
                  response.write "<tr><td HEIGHT=1 BGCOLOR=#CCCCCC width=700></td></tr>"
                  response.write "</Table>"
               end if
               if ObjRs("responsavel") <> "" then
                  Strq = "Select * from users where UserCode=" & ObjRs("responsavel")
                  set RstResp = ObjConn.execute(strq)
                  response.write "<table width=800 border=0 class=style9><tr><td>Veterinário responsável : " & RstResp("NomeCompleto") & " - CRMV : "& RstResp("CRM") &"</td></tr></table>"
                  rstResp.close
                  set RstResp = Nothing
               end if
               rstdetparts.close
               rstParts.close
               set rstDetparts = nothing
               set rstDetparts2 = nothing
               set rstParts = nothing
            else ' Aqui começa o exame em forma de laudo
                response.write "<table width=800 nowrap border=0 class=style9>"
                response.write "<tr>"
                strq = "select * from detalhesLaudosConsultas where CodExameConsulta =" & Id & " order by autocode"
                Set RstDetLau = Objconn.execute(strq)
                rstDetLau.movefirst
                if isnull(rstDetLau("titulo")) = false then response.write "<td><b>Exame : " & rstDetLau("Titulo") & "</td></tr>"
                response.write "<td width=800 div align=Left>"
                response.write replace((rstDetLau("valor")& ""),VbCrLf,"<br>")
                response.write "</td></tr></table>"

'                response.write "<table width=700 border=0 class=style9><tr><td>Exame liberado pelo funcionário " & ObjRs("Responsavel") & "</td></tr>"
'                response.write "<tr><td>Veterinário responsável : " & ObjRs("eletronicamente") & "</td></tr>"
                 
               If ObjRs("responsavel") <> "" then
                  Strq = "Select * from users where UserCode=" & ObjRs("responsavel")
                  set RstResp = ObjConn.execute(strq)
                  response.write "<table width=800 border=0 class=style9><tr><td>Veterinário responsável : " & RstResp("NomeCompleto") & " - CRMV : "& RstResp("CRM") &"</td></tr></table>"
                  rstResp.close
                  set RstResp = Nothing
               end if
                response.write "</td><tr><td colspan=3 height=1 bgcolor=#000000></td></tr>"
                Response.Write "</table>"
             End if
          end if
          CodEx=""
       End if
       x = x + 1
    loop
%>
</div>
<%
ObjConn.close
set rstExam = nothing
set ObjConn = nothing
End if
%>