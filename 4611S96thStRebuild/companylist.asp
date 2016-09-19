<!DOCTYPE HTML>
<html>
<head>
	<title>Directory Solutions - Company List</title>
	<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="No-Cache">
	<META HTTP-EQUIV="Expires" CONTENT="-1">
	<link rel="stylesheet" href="style_ds.css">
</head>

<body background="bg_list.jpg" bgproperties="fixed" leftmargin="0" topmargin="0" class="listingsWidth">
<div class="pushdown"></div>

	<%
	Session.timeout = 15
	Set conn = server.createObject("ADODB.Connection")
	conn.open "4611S96thSt"
	Set Session("MyDB_conn")=conn
	%>

	<%
	StrSQLQuery = "SELECT * FROM Companies"
	Set CompCount = Server.CreateObject("ADODB.recordset")
	CompCount.Open strSQLQuery, conn, 3, 3
	%>

	<%
	StrSQLQuery = "SELECT * FROM Companies WHERE CompanyName LIKE '"&request.querystring("letter")&"%'	ORDER BY CompanyName ASC"
	Set qCompanies = Server.CreateObject("ADODB.recordset")
	qCompanies.Open strSQLQuery, conn, 3, 3
	%>


<table id="tbl_listing" border="0" cellpadding="0" cellspacing="0" class="listingsWidth">

<% 'Put in this section a 'no matches' message if there are no listings with the selected letter
IF qCompanies.recordcount = 0 THEN%>
		<tr>
				<td align="center" valign="top" class="setOne0"></td><td align="left" colspan="5" class="setOne6"><div class="listLarge"><% IF (CompCount.recordcount = 0) AND (request.querystring("letter") = "") THEN%>There are no listings at this time<%ELSE%>There are no listings that start with "<%response.write (request.querystring("letter"))%>"<%END IF%></div></td>
   		</tr>
<%
ELSE
%>

<%'If there are records to display, then start displaying them. %>
<% do until qCompanies.EOF%>

<tr>
	<td align="center" valign="top" class="setOne0">
		<%
		DIM floor : suite = qCompanies("suite")
		char1 = Left(suite,1)
		IF char1 <> "3" THEN %>
		<a class="wayfinder" href="#" target="_top" data-company="<%response.write(qCompanies("CompanyID"))%>">
      <img src="btn_map.jpg" border="0">
    </a>
		<% END IF %>
	</td>
	<td align="left" valign="top" class="setOne1"><div class="listLarge"><%response.write (qCompanies("CompanyName"))%></div></td>
	<td align="center" valign="top" class="setOne2"><div class="listLarge"><%response.write (qCompanies("Suite"))%></div></td>
	<td align="center" valign="top" class="setOne3"><div class="listLarge"><%response.write (qCompanies("floor")) %></div></td><!-- elevator -->
	<td align="center" valign="top" class="setOne4"><div class="listLarge"></div></td>
	<td align="center" valign="top" class="setOne5"><div class="listLarge"></div></td>
</tr>

<% qCompanies.movenext
loop %>
<%END IF%>

<%'Next, put in some empty rows to allow for additional scrolling
'For count = 0 to 15
For count = 0 to 0%>
<tr>
	<td colspan="6" height="45"><div class="listLarge"></div></td>
</tr>
<% next%>

</table>


<% conn.close
set conn = nothing
%>
<% '/*
'*Move the scripts to the bottom of the body so they are non blocking
'*/ %>
<script language="javascript" src="scriptFiles/jquery.min.js"></script>
<script language="javascript" src="scriptFiles/jquery.cookie.js"></script>
<script language="javascript" src="scriptFiles/lightbox/lightbox.js"></script>
<script language="javascript" src="scriptFiles/ready.js"></script>
</body>
</html>
