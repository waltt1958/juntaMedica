<HTML>
<HEAD>


<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
<title>CONVERSOR ARCHIVOS JM</title>

</HEAD>

<body onload="maximizar()">

<H5>Hoy es: <%=weekdayname(weekday(date()))%>, <%=date%></H5>
<h1>COVNERSOR ARCHIVOS TELEGRAMA JM</h1>
<br>
<br>
<br>
<br><br>
<br><br>
<br>
<br>
<br>
<br>
<br>

<%

FileName= Session("nombreARC")
Response.Clear 
Response.ContentType="application / octet-stream"  
Response.AddHeader "content-disposition", "attachment; filename=" & FileName


Set objStream = Server.CreateObject("ADODB.Stream")
objStream.Type= 1
Response.CharSet = "UTF-8"
objStream.Open
objStream.LoadFromFile Server.MapPath(FileName)
Response.AddHeader "Content-Disposition", "attachment; filename=" & FileName
Response.ContentType = "application/octet-stream"
while not objStream.EOS
	Response.BinaryWrite objStream.Read(1024 * 64)
Wend

objStream.Close
Set stream= Nothing
Response.Flush
Response.End



%>

</body>

</HTML>