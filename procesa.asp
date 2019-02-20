<HTML>
<HEAD>


<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
<title>CONVERSOR ARCHIVOS JM</title>

</HEAD>

<body onload="maximizar()">

<!--#include virtual="/conectar.asp"-->

<H5>Hoy es: <%=weekdayname(weekday(date()))%>, <%=date%></H5>
<h1>CONVERSOR ARCHIVO TELEGRAMAS DE LA JUNTA MEDICA</h1>
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
if Session("carga")= 1 then

recupera= Session("archivo")
archivo= "c:\inetpub\wwwroot\juntaMedica\" & recupera

sqlLIMPIA = "DELETE * from sancor"
conectarOEP.execute sqlLIMPIA

sqlBORRA= "DELETE * from copiaSANCOR"
conectarOEP.execute sqlBORRA

Set objFSO = Server.CreateObject ("Scripting.FileSystemObject")

Set varArchivo = objFSO.OpenTextFile (archivo,1)

varArchivo.SkipLine

Do while not varArchivo.AtEndOfStream

	 arrayLinea = split (varArchivo.ReadLine, "|", - 1,1)

	sqlinsert= "INSERT INTO sancor (Apellido, Calle, CP, Localidad,Provincia, Operativa, Guia) VALUES ( '" & left(arrayLinea(0),30) & "','" & left(arrayLinea(1),30) & "','" & arrayLinea(2) & "', '"& left(arrayLinea(3),30) & "','" & left(arrayLinea(4),30) & "','" & arrayLinea(5) & "','" & arrayLinea(6) & "')"
	 
	conectarOEP.execute (sqlinsert)
 
loop 
	
Set varArchivo = Nothing
Set objFSO = Nothing

sqlINSERT="INSERT INTO copiaSANCOR select * from sancor"
conectarOEP.execute sqlINSERT

sqlACTUALIZA ="UPDATE copiaSANCOR SET copiaSANCOR.RETIdomicilio = 'Independencia', copiaSANCOR.RETInumero = '333', copiaSANCOR.RETIpiso ='0', copiaSANCOR.RETIdepto ='0', copiaSANCOR.RETIcp ='2322', copiaSANCOR.RETIlocalidad = 'Sunchales', copiaSANCOR.RETIprov = 'Santa Fe'"
conectarOEP.execute sqlACTUALIZA

 Set rsARCHIVO = Server.CreateObject("ADODB.recordset")

 sqlARCHIVO= "select * from copiaSANCOR"

 rsARCHIVO.open sqlARCHIVO, conectarOEP

 actual= now()

 nombre= "SANCOR " & day(actual) & "-" & month(actual) & "-" & year(actual) & "  "& hour(actual) & "-" & Minute(actual) & "-" & Second(actual) & ".txt"
 
  Set fso = Server.CreateObject ("Scripting.FileSystemObject")

  'Set arcTEXTO = fso.CreateTextFile(server.mappath("bajaSANCOR.txt"), true)
  Set arcTEXTO = fso.CreateTextFile(server.mappath(nombre), true)

  texto1 = rsARCHIVO.Fields(0).name & "|" & rsARCHIVO.Fields(7).name & "|" & rsARCHIVO.Fields(1).name & "|" & rsARCHIVO.Fields(8).name & "|" & rsARCHIVO.Fields(9).name & "|" & _
  rsARCHIVO.Fields(10).name & "|" & rsARCHIVO.Fields(2).name & "|" & rsARCHIVO.Fields(3).name & "|" & rsARCHIVO.Fields(4).name & "|" & rsARCHIVO.Fields(11).name & "|" & _
  rsARCHIVO.Fields(12).name & "|" & rsARCHIVO.Fields(13).name & "|" & rsARCHIVO.Fields(14).name & "|" & rsARCHIVO.Fields(15).name & "|" & rsARCHIVO.Fields(16).name _
  & "|" & rsARCHIVO.Fields(17).name & "|" & rsARCHIVO.Fields(18).name & "|" & rsARCHIVO.Fields(19).name & "|" & rsARCHIVO.Fields(20).name & "|" & _
  rsARCHIVO.Fields(21).name & "|" & rsARCHIVO.Fields(22).name & "|" & rsARCHIVO.Fields(23).name & "|" & rsARCHIVO.Fields(24).name & "|" & _
  rsARCHIVO.Fields(25).name & "|" & rsARCHIVO.Fields(26).name & "|" & rsARCHIVO.Fields(27).name & "|" & rsARCHIVO.Fields(5).name & "|" & rsARCHIVO.Fields(28).name _
  & "|" & rsARCHIVO.Fields(6).name & "|" & rsARCHIVO.Fields(29).name & "|" & rsARCHIVO.Fields(30).name & "|" & rsARCHIVO.Fields(31).name
  
  arcTEXTO.WriteLine(texto1)
 
  do while not rsARCHIVO.EOF

  texto= rsARCHIVO.Fields("Apellido") & "|" & rsARCHIVO("DESTnombre") & "|" & rsARCHIVO("Calle") & "|" & rsARCHIVO("DESTnumero") & "|" & rsARCHIVO("DESTpiso") & "|" & _
  rsARCHIVO("DESTdepto")  & "|" & rsARCHIVO("CP") & "|" & rsARCHIVO("Localidad") & "|" & rsARCHIVO("Provincia") & "|" & rsARCHIVO("DESTtelefono") & "|" & _
  rsARCHIVO("DESTemail") & "|" & rsARCHIVO("RETIdomicilio") & "|" & rsARCHIVO("RETInumero") & "|" & rsARCHIVO("RETIpiso") & "|" & rsARCHIVO("RETIdepto") _
  & "|" & rsARCHIVO("RETItelefono") & "|" & rsARCHIVO("RETIcp") & "|" & rsARCHIVO("RETIlocalidad") & "|" & rsARCHIVO("RETIprov") & "|" & rsARCHIVO("RETIcontacto") _
  & "|" & rsARCHIVO("PAQpeso") & "|" & rsARCHIVO("PAQalto") & "|" & rsARCHIVO("PAQalto") & "|" & rsARCHIVO("PAQancho") & "|" & rsARCHIVO("PAQvalor") _
  & "|" & rsARCHIVO("NROremito") & "|" & rsARCHIVO("Operativa") & "|" & rsARCHIVO("IMPremito") & "|" & rsARCHIVO("Guia") & "|" & rsARCHIVO("NROproducto") _
  & "|" & rsARCHIVO("RETIemail") & "|" & rsARCHIVO("observaciones")

  arcTEXTO.WriteLine(texto)

  rsARCHIVO.MoveNext

  loop

 rsARCHIVO.close
 Set rsARCHIVO= nothing
	
  Set fso = nothing
  Set arcTEXTO = nothing

sqlLIMPIA = "DELETE * from sancor"
conectarOEP.execute sqlLIMPIA

sqlBORRA= "DELETE * from copiaSANCOR"
conectarOEP.execute sqlBORRA

Session("nombreARC")= nombre

%>

<!--#include virtual="/desconectar.asp"-->

<%

else


response.redirect ("index.asp")

end if

%>
<table align="center">
<tr>
<td>
<a href="bajaArchivo.asp" target="_self"><input type="button" name="descarga" value="DESCARGAR ARCHIVO" style="FONT-SIZE: 20pt; border: 5px solid; [b]FONT-FAMILY: Verdana, boldt[/b];
BACKGROUND-COLOR: #C0C0C0"></a>
</td>
</tr>
</table>
</script>

<SCRIPT Language="javascript" type="text/javascript">

function maximizar() {

window.moveTo(0,0);

window.resizeTo(screen.width,screen.height);
}
</SCRIPT>


</body>

</HTML>