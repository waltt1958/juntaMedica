<HTML>
<HEAD>



<title>CONVERSOR ARCHIVOS JM</title>

</HEAD>

<body onload="maximizar()">

<!--#include virtual="/conectar.asp"-->

<table align="center">
<tr>
<td>
<a href="bajaArchivo.asp" target="_self"><input type="button" name="descarga" value="DESCARGAR ARCHIVO" style="FONT-SIZE: 20pt; border: 5px solid; [b]FONT-FAMILY: Verdana, boldt[/b];
BACKGROUND-COLOR: #C0C0C0"></a>
</td>
</tr>
</table>


<%



sqlLIMPIA = "DELETE * from datosJM"
conectarOEP.execute sqlLIMPIA

     Set cn = CreateObject("ADODB.Connection") 
     Set rs = CreateObject("ADODB.Recordset") 

     strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= c:\inetpub\wwwroot\juntaMedica\CITACIONES MARZO 2019.xls;Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'" 
					
     cn.Open strConnect 
    strSql = "SELECT * FROM [Hoja1$] " 
    rs.Open strSql, cn 

		 Do while not (isnull(rs(0)))
	
		conectarOEP.execute "INSERT INTO datosJM (orden, FechaEnvio, FechaJM, Hora, Agente, DNI, Direccion, Localidad, CP, lugarPRESENTACION, AREA, DomicilioPresentacion, Provincia) VALUES ('"&rs(0)&"','"&rs(1)&"','"&rs(2)&"','"&rs(3)&"','"&rs(4)&"','"&rs(5)&"','"&rs(6)&"','"&rs(7)&"','"&rs(8)&"','"&rs(9)&"','"&rs(10)&"','"&rs(11)&"','"&rs(12)&"')"
		' sqlINSERT = "INSERT INTO junta (Nº, Fecha envío, Fecha JM, Hora, Agente, DNI, Dirección, Localidad, CP, LUGAR PRESENTACION, AREA, Domicilio presentación, Provincia) VALUES ('" & rs("0") & "', '" & rs("1") & "','" & rs("2") & "','" & rs("3") & "', '" & rs("4") & "','" & rs("5") & "','" & rs("6") & "', '" & rs("7") & "','" & rs("8") & "','" & rs("9") & "', '" & rs("10") & "','" & rs("11") & "','" & rs("12") & "')"
				
		 rs.MoveNext 
      Loop 
	  
	
	
	
	
	

%>
<!--#include virtual="/desconectar.asp"-->




</body>

</HTML>