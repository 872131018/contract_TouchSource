<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>Directory Solutions - List</title>
	<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="No-Cache">
	<META HTTP-EQUIV="Expires" CONTENT="-1">
	<link rel="stylesheet" href="style_ds.css">
</head>

<body background="bg_list.jpg" bgproperties="fixed" leftmargin="0" topmargin="0">
<div class="listXS">&nbsp;</div>

	<%
	Session.timeout = 15
	Set conn = server.createObject("ADODB.Connection")
	conn.open "DalhousieUnivMedicine"
	Set Session("MyDB_conn")=conn
	%>
	
	<%
	StrSQLQuery = "SELECT *	FROM List12  WHERE (Text2 = 'M') OR (Text2 = 'E') ORDER BY Text1 ASC"
	Set List12List = Server.CreateObject("ADODB.recordset")
	List12List.Open strSQLQuery, conn, 3, 3
	%>

<table border="0" cellpadding="0" cellspacing="0" width="101%">

<% 'Put in this section a 'no matches' message if there are no listings with the selected letter
IF List12List.recordcount = 0 THEN%>
			<tr>
				<TD align="center" valign="top" class="setSix0"></td>
				<TD align="left" valign="middle" colspan="5" class="setSix6"><div class="listMed">There are no listings</div></TD>
			</TR>
<%
ELSE
%>

<%'If there are records to display, then start displaying them. %>
<% do until List12List.EOF%>
<TR>
	<TD align="center" valign="top" class="setSix0">	<a href="list12animatedmap.asp?ID=<%response.write(List12List("ID"))%>" target="_top"><img src="btn_map.jpg" border="0"></a></td>
	<TD align="left" valign="top" class="setSix1"><div class="listMed"><%response.write (List12List("Text1"))%></div></td>
	<TD align="left" valign="top" class="setSix2"><div class="listMed"><%response.write (List12List("Text3"))%></div></td>
	<TD align="left" valign="top" class="setSix3"><div class="listMed"><%response.write (List12List("Text7"))%></div></td>
	<TD align="left" valign="top" class="setSix4"><div class="listMed"><%response.write (List12List("Text4"))%></div></TD>
	<TD align="left" valign="top" class="setSix5"><div class="listMed"></div></TD>
</tr>

<% List12List.movenext
loop %>
<%END IF%>

<%'Next, put in some empty rows to allow for additional scrolling
For count = 0 to 15%>
<TR>
	<TD colspan="5" height="45"><div class="listLarge"></div></TD>
</tr>
<% next%>

</TABLE>

<% conn.close
set conn = nothing
%>

</body>
</html>
