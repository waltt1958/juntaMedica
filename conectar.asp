<%


set conectarOEP = Server.CreateObject("ADODB.Connection")
' conectarOEP.ConnectionString = "DBQ=C:\inetpub\wwwroot\juntaMedica\JuntaMedica.mdb;DRIVER={MS Access (*.mdb)}" 
' conectarOEP.Open 
conectarOEP.Open = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("JuntaMedica.mdb")


%>