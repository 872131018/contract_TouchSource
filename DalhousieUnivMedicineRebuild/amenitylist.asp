<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>Directory Solutions - Amenity List</title>
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
	StrSQLQuery = "SELECT *	FROM Amenities ORDER BY AmenityName ASC"
	Set Amenities = Server.CreateObject("ADODB.recordset")
	Amenities.Open strSQLQuery, conn, 3, 3
	%>

<table border="0" cellpadding="0" cellspacing="0" width="101%">

<% 'Put in this section a 'no matches' message if there are no listings with the selected letter
IF Amenities.recordcount = 0 THEN%>
			<tr>
				<TD align="right" valign="top" class="setThree0"></td>
				<TD align="left" valign="middle" colspan="5" class="setThree6"><div class="listMed">There are no listings</div></TD>
			</TR>
<%
ELSE
%>

<%'If there are records to display, then start displaying them. %>
<% do until Amenities.EOF%>

<TR>
	<TD align="center" valign="top" class="setThree0"><a href="amenitiesanimatedmap.asp?ID=<%response.write(Amenities("AmenityID"))%>" target="_top"><img src="btn_map.jpg" border="0"></a></td>
	<TD align="left" valign="top" class="setThree1"><div class="listMed"><%response.write (Amenities("AmenityName"))%></div></td>
	<TD align="left" valign="top" class="setThree2"><div class="listMed"><%response.write (Amenities("Location"))%></div></td>
	<TD align="left" valign="top" class="setThree3"><div class="listMed"><%response.write (Amenities("Building"))%></div></td>
	<TD align="center" valign="top" class="setThree4"><div class="listMed"></div></TD>
	<TD align="center" valign="top" class="setThree5"><div class="listMed"></div></TD>
</tr>

<% Amenities.movenext
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
