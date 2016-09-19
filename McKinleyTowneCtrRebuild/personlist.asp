<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>Directory Solutions - Person List</title>
	<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="No-Cache">
	<META HTTP-EQUIV="Expires" CONTENT="-1">
	<link rel="stylesheet" href="style_ds.css">
</head>

<body background="bg_list.jpg" bgproperties="fixed" leftmargin="0" topmargin="0">
<div class="listXS">&nbsp;</div>

	<%
	Session.timeout = 15
	Set conn = server.createObject("ADODB.Connection")
	conn.open "McKinleyTowneCtr"
	Set Session("MyDB_conn")=conn
	%>
	
	<%
	StrSQLQuery = "SELECT * FROM Individuals"
	Set IndivCount = Server.CreateObject("ADODB.recordset")
	IndivCount.Open strSQLQuery, conn, 3, 3
	%>
	
	<%
	StrSQLQuery = "SELECT Individuals.IndividualName, Individuals.Title, Companies.CompanyName, Individuals.Suite, Individuals.Phone, Individuals.CompanyID, Individuals.IndID, Companies.CompanyID, Companies.Floor, Companies.Building	FROM Companies 	INNER JOIN Individuals ON Companies.CompanyID = Individuals.CompanyID	WHERE Individuals.IndividualName LIKE '"&request.querystring("letter")&"%'	ORDER BY Individuals.IndividualName ASC"
	Set Names = Server.CreateObject("ADODB.recordset")
	Names.Open strSQLQuery, conn, 3, 3
	%>

<table border="0" cellpadding="0" cellspacing="0" width="101%">
<% 'Put in this section a 'no matches' message if there are no listings with the selected letter
IF Names.recordcount = 0 THEN%>
		<tr>
				<TD colspan="5" height="50"><div class="listMed">&nbsp;&nbsp;<% IF (IndivCount.recordcount = 0) AND (request.querystring("letter") = "") THEN%>There are no listings at this time.<%ELSE%>There are no listings that start with "<%response.write (request.querystring("letter"))%>".<%END IF%></div></TD>
		</TR>
<%
ELSE
%>

<%'If there are records to display, then start displaying them. %>
<% do until Names.EOF%>
<a href="companydetail.asp?ID=<%response.write(Names("CompanyID"))%>" target="_top">
<tr>
	<TD align="left" valign="top" class="setTwo0"><div class="listMed"><a href="companyanimatedmap.asp?ID=<%response.write(Names("CompanyID"))%>" target="_top"><img src="btn_map.jpg" border="0"></a></div></TD>
	<TD align="left" valign="top" class="setTwo1"><div class="listMed"><%response.write (Names("IndividualName"))%></div><%IF Names("Title") > "" Then%><div class="ListSmall">&nbsp;&nbsp; - <%response.write (Names("Title"))%><%END IF%></div></td>
	<TD align="left" valign="top" class="setTwo2"><div class="listMed"><%response.write (Names("CompanyName"))%></div></td>
	<TD align="center" valign="top" class="setTwo3"><div class="listMed"><%response.write (Names("Phone"))%></div></td>
	<TD align="center" valign="top" class="setTwo4"><div class="listMed"></div></TD>
	<TD align="center" valign="top" class="setTwo5"><div class="listMed"></div></TD>
</tr>
</a>
<% Names.movenext
loop %>
<%END IF%>

<%'Next, put in some empty rows to allow for additional scrolling
For count = 0 to 15%>
		<tr>
			<TD colspan="5" height="45"><div class="listMed"></div></TD>
		</TR>
<% next%>

</TABLE>

<% conn.close
set conn = nothing
%>

</body>
</html>
