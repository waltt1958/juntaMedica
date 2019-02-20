<HTML>
<HEAD>

<meta charset="utf-8">
<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
<title>CONVERSOR ARCHIVOS RAF</title>

</HEAD>

<body onload="maximizar()">
<br>

<H5>Hoy es: <%=weekdayname(weekday(date()))%>, <%=date%></H5>
<h1>SUC. OCA RAFAELA - PAQUETERIA (Oper. 288140 )</h1>
<br>
<hr size= 6 color="black"></hr>

<br>

<h3>Recuerde que el archivo que recibió de Sancor lo debe guardar en esta PC. Le será solicitada la ubicación donde lo grabó durante el proceso de conversión</h3>

<br>

<hr size= 6 color="black"></hr>

<br>
<br>

<table align="center">
<tr align="center"><td><input type="button" class="button" name="iniciar" onclick=location.href='cargaArchivo.asp' value="     INICIAR PROCESO     "></td></tr>
</table>

<%
texto= "txt"
bbdd= "mdb"
clasico= "asp"
forma= "css"
imagen = "png"

Set objFSO = server.CreateObject ("Scripting.FileSystemObject")
set objFolder=objFSO.GetFolder("c:\inetpub\wwwroot\juntaMedica\")

for each objFile in objFolder.files

Select case objFSO.GetExtensionName(objFile)
case bbdd
case clasico
case forma
case imagen
case else

objFile.delete
end select

next

Session("inicio")= 1
%>
</script>

<SCRIPT Language="javascript" type="text/javascript">

function maximizar() {

window.moveTo(0,0);

window.resizeTo(screen.width,screen.height);
}
</SCRIPT>


</body>

</HTML>