<HTML>
<HEAD>


<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
<title>CONVERSOR ARCHIVOS JM</title>

</HEAD>

<body onload="maximizar()">

<!--#include virtual="/conectar.asp"-->

<H5>Hoy es: <%=weekdayname(weekday(date()))%>, <%=date%></H5>
<h1>CONVERSOR ARCHIVO TELEGRAMAS DE LA JM</h1>
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

sqlLIMPIA = "DELETE * from datosJM"
conectarOEP.execute sqlLIMPIA

     Set cn = CreateObject("ADODB.Connection") 
     Set rs = CreateObject("ADODB.Recordset") 

     strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & archivo & " ;Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'" 
					
     cn.Open strConnect 
    strSql = "SELECT * FROM [Hoja1$] " 
    rs.Open strSql, cn 
	
	 Do while not rs.EOF 
	
		conectarOEP.execute "INSERT INTO datosJM (orden, FechaEnvio, FechaJM, Hora, Agente, DNI, Direccion, Localidad, CP, lugarPRESENTACION, AREA, DomicilioPresentacion, Provincia) VALUES ('"&rs(0)&"','"&rs(1)&"','"&rs(2)&"','"&rs(3)&"','"&rs(4)&"','"&rs(5)&"','"&rs(6)&"','"&rs(7)&"','"&rs(8)&"','"&rs(9)&"','"&rs(10)&"','"&rs(11)&"','"&rs(12)&"')"
		' sqlINSERT = "INSERT INTO junta (Nº, Fecha envío, Fecha JM, Hora, Agente, DNI, Dirección, Localidad, CP, LUGAR PRESENTACION, AREA, Domicilio presentación, Provincia) VALUES ('" & rs("0") & "', '" & rs("1") & "','" & rs("2") & "','" & rs("3") & "', '" & rs("4") & "','" & rs("5") & "','" & rs("6") & "', '" & rs("7") & "','" & rs("8") & "','" & rs("9") & "', '" & rs("10") & "','" & rs("11") & "','" & rs("12") & "')"
				
		 rs.MoveNext 
      Loop 

set	rs= nothing
cn.close
set cn = nothing






 ' Set rsARCHIVO = Server.CreateObject("ADODB.recordset")

 ' sqlARCHIVO= "select * from juntaMedica"

 ' rsARCHIVO.open sqlARCHIVO, conectarOEP

 ' actual= now()

 ' nombre= "JM " & day(actual) & "-" & month(actual) & "-" & year(actual) & "  "& hour(actual) & "-" & Minute(actual) & "-" & Second(actual) & ".txt"
 
  ' Set fso = Server.CreateObject ("Scripting.FileSystemObject")

  'Set arcTEXTO = fso.CreateTextFile(server.mappath("bajaSANCOR.txt"), true)
  ' Set arcTEXTO = fso.CreateTextFile(server.mappath(nombre), true)

   ' do while not rsARCHIVO.EOF

  ' texto= rsARCHIVO.Fields("Apellido") & "|" & rsARCHIVO("DESTnombre") & "|" & rsARCHIVO("Calle") & "|" & rsARCHIVO("DESTnumero") & "|" & rsARCHIVO("DESTpiso") & "|" & _
  ' rsARCHIVO("DESTdepto")  & "|" & rsARCHIVO("CP") & "|" & rsARCHIVO("Localidad") & "|" & rsARCHIVO("Provincia") & "|" & rsARCHIVO("DESTtelefono") & "|" & _
  ' rsARCHIVO("DESTemail") & "|" & rsARCHIVO("RETIdomicilio") & "|" & rsARCHIVO("RETInumero") & "|" & rsARCHIVO("RETIpiso") & "|" & rsARCHIVO("RETIdepto") _
  ' & "|" & rsARCHIVO("RETItelefono") & "|" & rsARCHIVO("RETIcp") & "|" & rsARCHIVO("RETIlocalidad") & "|" & rsARCHIVO("RETIprov") & "|" & rsARCHIVO("RETIcontacto") _
  ' & "|" & rsARCHIVO("PAQpeso") & "|" & rsARCHIVO("PAQalto") & "|" & rsARCHIVO("PAQalto") & "|" & rsARCHIVO("PAQancho") & "|" & rsARCHIVO("PAQvalor") _
  ' & "|" & rsARCHIVO("NROremito") & "|" & rsARCHIVO("Operativa") & "|" & rsARCHIVO("IMPremito") & "|" & rsARCHIVO("Guia") & "|" & rsARCHIVO("NROproducto") _
  ' & "|" & rsARCHIVO("RETIemail") & "|" & rsARCHIVO("observaciones")

  ' arcTEXTO.WriteLine(texto)

  ' rsARCHIVO.MoveNext

  ' loop

 ' rsARCHIVO.close
 ' Set rsARCHIVO= nothing
	
  ' Set fso = nothing
  ' Set arcTEXTO = nothing

' sqlLIMPIA = "DELETE * from juntaMedica"
' conectarOEP.execute sqlLIMPIA


' Session("nombreARC")= nombre

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