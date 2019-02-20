<html>

<head>

<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
</head>

<body>


<h1>SUC. OCA RAFAELA - PAQUETERIA (Oper. 288140 )</h1>


<%@LANGUAGE="VBSCRIPT"%> 

<%
if Session("inicio") = 1 then

Session("carga")= 1

response.buffer=true
Func = Request("Func")
if isempty(Func) Then
Func = 1
End if
Select Case Func
Case 1

%>

<table width="650" border="0" align="center">
  <tr>
    <td>
    <div align="center">
      <h3>Seleccione el archivo que quiere procesar</h3>
    </div>
    </td>
  </tr>
</table>
<form enctype="multipart/form-data" action="cargaArchivo.asp?func=2" method="POST" id="form1" name="form1">
  <table align="center">
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td><h5>Presione el botón "SELECCIONAR ARCHIVO" y elija el archivo que quiere procesar.</h5> <br>
            <br>
      </font></td>
    </tr>
    <tr>
      <td><font color="black" size="4"><b>Luego pulsa el botón subir.</b><br>
      <br>
      </font></td>
    </tr>
    <tr>
      <td><strong><font color="black" size="4">Nombre del archivo...</font></strong></td>
    </tr>
    <tr>
      <td><font size="2"><input name="File1" size="100" accept=".txt" type="file"> 
      </font></td>
    </tr>
    <tr>
      <td align="left"><input type="submit" value="Subir"> <br>
      <br>
      </td>
    </tr>
    <tr>
      <td><font color="black" size="4"><b>NOTA:</b> Espere, recibirá una notificación cuando el archivo haya sido subido</font><font size="4">.<br>
      <br>
      </font></td>
    </tr>
  </table>
  
  
  <% 'Código ASP

Case 2

ForWriting = 2
adLongVarChar = 201
lngNumberUploaded = 0

'Get binary data from form 
noBytes = Request.TotalBytes 
binData = Request.BinaryRead (noBytes)

'convery the binary data to a string
Set RST = CreateObject("ADODB.Recordset")
LenBinary = LenB(binData)

if LenBinary > 0 Then
	RST.Fields.Append "myBinary", adLongVarChar, LenBinary
	RST.Open
	RST.AddNew
	RST("myBinary").AppendChunk BinData
	RST.Update
	strDataWhole = RST("myBinary")
End if

strBoundry = Request.ServerVariables ("HTTP_CONTENT_TYPE")
lngBoundryPos = instr(1,strBoundry,"boundary=") + 8 
strBoundry = "--" & right(strBoundry,len(strBoundry)-lngBoundryPos)

'Get first file boundry positions.
lngCurrentBegin = instr(1,strDataWhole,strBoundry)
lngCurrentEnd = instr(lngCurrentBegin + 1,strDataWhole,strBoundry) - 1

Do While lngCurrentEnd > 0
	'Get the data between current boundry and remove it from the whole.
	strData = mid(strDataWhole,lngCurrentBegin, lngCurrentEnd - lngCurrentBegin)
	strDataWhole = replace(strDataWhole,strData,"")

	'Get the full path of the current file.
	lngBeginFileName = instr(1,strdata,"filename=") + 10
	lngEndFileName = instr(lngBeginFileName,strData,chr(34)) 
	'Make sure they selected at least one file. 
	if lngBeginFileName = lngEndFileName and lngNumberUploaded = 0 Then

		Response.Write "<H2> Ha ocurrido el siguiente error.</H2>"
		Response.Write "Debe elegir un archivo para subir"
		Response.Write "<BR><BR>Pulse el botón VOLVER y realice la corrección."
		Response.Write "<BR><BR><INPUT type='button' onclick='history.go(-1)' value='<< Volver' id='button'1 name='button'1>"
		Response.End 
	End if

'There could be one or more empty file boxes. 

if lngBeginFileName <> lngEndFileName Then
	strFilename = mid(strData,lngBeginFileName,lngEndFileName - lngBeginFileName)


	'Loose the path information and keep just the file name. 
	tmpLng = instr(1,strFilename,"\")
	Do While tmpLng > 0
		PrevPos = tmpLng
		tmpLng = instr(PrevPos + 1,strFilename,"\")
	Loop

	FileName = right(strFilename,len(strFileName) - PrevPos)

	'Get the begining position of the file data sent.
	'if the file type is registered with thebrowser then there will be a Content-Type
	lngCT = instr(1,strData,"Content-Type:")

	if lngCT > 0 Then
		lngBeginPos = instr(lngCT,strData,chr(13) & chr(10)) + 4
	Else
		lngBeginPos = lngEndFileName
	End if
	'Get the ending position of the file dat
	' a sent.
	lngEndPos = len(strData) 

	'Calculate the file size. 
	lngDataLenth = lngEndPos - lngBeginPos
	'Get the file data 
	strFileData = mid(strData,lngBeginPos,lngDataLenth)

	'Create the file. 
	Set fso = CreateObject("Scripting.FileSystemObject")

	'Lo guarda en la carpeta actual
	Set f = fso.OpenTextFile(server.mappath(".\") & "/" & FileName, ForWriting, True)
	f.Write strFileData
	Set f = nothing
	Set fso = nothing


	lngNumberUploaded = lngNumberUploaded + 1

	End if

	'Get then next boundry postitions if any.
	lngCurrentBegin = instr(1,strDataWhole,strBoundry)
	lngCurrentEnd = instr(lngCurrentBegin + 1,strDataWhole,strBoundry) - 1
loop

%>
<div align="center">
<br>
<br>
<br>
<br>
<font size="4">
<b>
<%
Session("archivo")= FileName
Response.Write "Archivo subido<Br>"
Response.Write lngNumberUploaded & " archivo ya está en el servidor.<BR>"
Response.Write "<BR><BR><INPUT type='button' class='button' onclick='document.location=" & chr(34) & "procesa.asp" & chr(34) & "' value='<< Continuar proceso' id='button'1 name='button'1>" 
%>
</b>
</font>
</div>
<%
End Select 

else	

Session.Contents.Remove("inicio")
response.redirect ("index.asp")

end if

%></form>

</body>

</html>
