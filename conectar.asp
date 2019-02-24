<%
set conectarOEP = Server.CreateObject("ADODB.Connection")
response.write ("antesssssssssssssssssssssssssssssssssssssss")
conectarOEP.Open = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("juntaMedica.mdb")
response.write ("despuexxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
%>