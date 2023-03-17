<style TYPE ="text/css">
A{text-decoration:none;font:Verdana;}
p{font-family:Arial;color:#000000;font-size:100%;}
.style9{font-size:9px;font-family:Verdana;color=#000000;}
#LayerTopo{position:absolute;width:798px;height:300px;z-index:1;Left:0;top:0px}
#Layer1{position:absolute;width:798px;height:300px;z-index:1;Left:10;top:65px;overflow=scroll}
#LayerBarraDeBaixo{position:absolute;width:700px;height:20px;z-index:1;Left:8;top:382px}
#LayerBarraDeCima{position:absolute;width:700px;height:20px;z-index:1;Left:8;top:80px}
#LayerBtnBack{position:absolute;width:300px;height:20px;z-index:10;Left:680;top:80px}
#LayerTxtBack{position:absolute;width:200px;height:20px;z-index:10;Left:600;top:81px}
</style><html><head><title>Sistema Veterinário - Central de Exames</title><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
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

   response.write "<div id=Topo><table width=780 class=style9 border=0>"
   response.write "<tr><td bgcolor=#f9f9f9 width=470 div align=left><img src=../scr/images/LogoNovo450.jpg border=0></img></td>"
   response.write "<td bgcolor=#f9f9f9 width=20 align=right><a href=AnimalInterface.asp?id="&id&"&origem="&origem&"&CC="&RstCadastro("codcli")&"><img src=../scr/images/SetaAtras.png border=0></img></a></td>"
   response.write "<td width=100 bgcolor=#f9f9f9><a href=AnimalInterface.asp?id="&id&"&origem="&origem&"&CC="&RstCadastro("codcli")&"><b>Voltar à lista.</a></td></tr>"
   response.write "</table></div>"

   response.write "<div id=layer1><table width=780 class=style9 border=0>"
   response.write "<tr><td colspan=2 div align=center><hr><B>Resultado de exames veterinários.</b></td></tr>"
   response.write "<tr><td bgcolor=#f7f7f7 colspan=2 div align=center>Pet selecionado : "& RstCadastro("NomeAnimal") &"</b></td></tr>"
   response.write "<tr><td colspan=2 bgcolor=#cccccc height=2></td></tr>"
   iF RstExames.eof = false then
      response.write "<tr><td bgcolor=#ccccFF height=16  colspan=2 div align=Left><B>Selecione um exame :</td></tr>"
      Do Until RstExames.eof = true
         if bgcolor="f7f7f7" then bgcolor="ffffff" else bgcolor="f7f7f7"
         Databr   = RIGHT("0" & day(RstExames("data")),2) & "/" & RIGHT("0" & month(RstExames("data")),2) & "/" & RIGHT("0" & year(RstExames("data")),2)
         response.write "<tr><td bgcolor=#"& bgcolor &" colspan=2 div align=Left><img src=../scr/images/lilstar.png border=0></img><a href=PrintExamNewONLINE.asp?CodCli=" & RstCadastro("codcli") & "&text=|"& RstExames("CodExameConsulta") &"| target=_self><font color=#000000> "& DataBr & " - " & RstExames("Nomexame") &"</a></td></tr>"
      RstExames.movenext
      loop
   end if
   response.write "</table></div>"
'   response.write "<tr><td height=18 background=../graphics/BkgGray.jpg colspan=2 div align=right></td></tr>"
'   response.write "<tr><td height=50 colspan=2 bgcolor=#f9f9f9></td></tr>"
'   response.write "</table>"
   RstExames.close
   Set RstExames = nothing
'   response.write "<table border=0 width=800 class=style9>"
else
   str ="animalInterface.asp?msg=Não existem exames concluídos para o pet selecionado.&CC=" & Session("IdCl")
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
response.write "<Tr><td colspan=2 bgcolor=#f9f9f9 height=50><hr></td></tr>"
response.write "</table></div>"



ObjConn.close
set RstCliente = nothing
Set ObjConn = nothing
Set SConnString = nothing
Set Strq = nothing
%>