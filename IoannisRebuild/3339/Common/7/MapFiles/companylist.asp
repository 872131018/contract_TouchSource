<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>Directory Solutions - Company List</title>
	<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="No-Cache">
	<META HTTP-EQUIV="Expires" CONTENT="-1">
	<link rel="stylesheet" href="style_ds.css">
</head>

<body background="bg_list.jpg" bgproperties="fixed" leftmargin="0" topmargin="0">
<div class="listXS">&nbsp;</div>

	<%
	Session.timeout = 15
	Set conn = server.createObject("ADODB.Connection")
	conn.open "Update"
	Set Session("MyDB_conn")=conn
	%>
	
	<%
	StrSQLQuery = "SELECT * FROM Companies WHERE (CompanyName LIKE '"&request.querystring("letter")&"%') AND (PropNum  = "&request.querystring("PropNum")&")	ORDER BY CompanyName ASC"
	Set Companies = Server.CreateObject("ADODB.recordset")
	Companies.Open strSQLQuery, conn, 3, 3
	%>

<table border="0" cellpadding="0" cellspacing="0" width="101%">

<% 'Put in this section a 'no matches' message if there are no listings with the selected letter
IF Companies.recordcount = 0 THEN%>
		<tr>
				<TD colspan="5" height="50"><div class="listLarge">&nbsp;&nbsp;There are no listings that start with this letter</div></TD>
   		</TR>
<%
ELSE
%>

<%'If there are records to display, then start displaying them. %>
<% do until Companies.EOF%>
<a href="companydetail.asp?ID=<%response.write(Companies("CompanyID"))%>" target="_top">
<TR>
	<TD align="left" valign="middle" class="setOne1"><div class="listLarge"><%response.write (Companies("CompanyName"))%></div></td>
	<TD align="center" valign="middle" class="setOne2"><div class="listLarge"><%response.write (Companies("Suite"))%></div></td>
	<TD align="center" valign="middle" class="setOne3"><div class="listLarge"><a href="companyanimatedmap.asp?ID=<%response.write(Companies("CompanyID"))%>" target="_top"><img src="btn_map.jpg" border="0"></a></div></td>
	<TD align="center" valign="middle" class="setOne4"><div class="listLarge"></div></TD>
	<TD align="center" valign="middle" class="setOne5"><div class="listLarge"></div></TD>
</tr>
</a>
<% Companies.movenext
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
