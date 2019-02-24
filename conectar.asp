<%
on error resume next
set conectarOEP = Server.CreateObject("ADODB.Connection")
conectarOEP.ConnectionString = "DBQ=C:\inetpub\wwwroot\juntaMedica\JuntaMedica.mdb;DRIVER={MS Access (*.mdb)}" 
conectarOEP.Open '= "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("JuntaMedica.mdb")
response.write("anduvo")
if err.number > 0 then 
response.write(err.description)
end if
%>