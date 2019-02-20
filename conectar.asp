<%
set conectarOEP = Server.CreateObject("ADODB.Connection")
conectarOEP.Open = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("sancor.mdb")
%>