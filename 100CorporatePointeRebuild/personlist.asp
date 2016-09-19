<!DOCTYPE HTML>
<html>
<head>
	<title>Directory Solutions - Person List</title>
	<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="No-Cache">
	<META HTTP-EQUIV="Expires" CONTENT="-1">
	<link rel="stylesheet" href="style_ds.css">
</head>

<body background="bg_list.jpg" bgproperties="fixed" leftmargin="0" topmargin="0" class="listingsWidth">
<div class="pushdown"></div>

	<%
	Session.timeout = 15
	Set conn = server.createObject("ADODB.Connection")
	conn.open "100CorporatePointe"
	Set Session("MyDB_conn")=conn
	%>
	
	<%
	StrSQLQuery = "SELECT * FROM Individuals"
	Set IndivCount = Server.CreateObject("ADODB.recordset")
	IndivCount.Open strSQLQuery, conn, 3, 3
	%>
	
	<%
	StrSQLQuery = "SELECT Individuals.IndividualName, Individuals.Title, Companies.CompanyName, Individuals.Suite, Individuals.iBuilding, Individuals.Phone, Individuals.CompanyID, Individuals.IndID, Individuals.IndImage, Individuals.IndCustom1, Individuals.IndCustom2, Companies.CompanyID, Companies.Floor, Companies.Building	FROM Companies 	INNER JOIN Individuals ON Companies.CompanyID = Individuals.CompanyID	WHERE Individuals.IndividualName LIKE '"&request.querystring("letter")&"%'	ORDER BY Individuals.IndividualName ASC"
	Set Names = Server.CreateObject("ADODB.recordset")
	Names.Open strSQLQuery, conn, 3, 3
	%>


<table id="tbl_listing" border="0" cellpadding="0" cellspacing="0" class="listingsWidth">

<% 'Put in this section a 'no matches' message if there are no listings with the selected letter
IF Names.recordcount = 0 THEN%>
		<tr>
				<td align="center" valign="top" class="setTwo0"></td><td align="left" colspan="5" class="setTwo6"><div class="listMed"><% IF (Names.recordcount = 0) AND (request.querystring("letter") = "") THEN%>There are no listings at this time<%ELSE%>There are no listings that start with "<%response.write (request.querystring("letter"))%>"<%END IF%></div></td>
		</TR>
<%
ELSE
%>

<%'If there are records to display, then start displaying them. %>
<% do until Names.EOF%>

<tr>
	<td align="center" valign="top" class="setTwo0">
<!-- Check to see if there is a jpg or pdf, set href to point to those pages, otherwise, it goes to the detail page via the link icon. -->
	<%IF Names("IndImage") > "" Then%>
		<% DIM IndImageName %>
		<% IndImageName = (Names("IndImage"))%>
		<% DIM IndImageLast3 %>
		<% IndImageLast3 = right(IndImageName, 3) %>
		<%IF IndImageLast3 = "pdf" Then%>
			<a href="indPDFDisplay.asp?ID=<%response.write(Names("IndID"))%>" target="_top">
		<%ELSEIF IndImageLast3 = "jpg" Then%>
			<a href="indJPGDisplay.asp?ID=<%response.write(Names("IndID"))%>" target="_top">
		<%END IF%>
	<%ELSE%>
		<a href="companydetail.asp?ID=<%response.write(Names("CompanyID"))%>" target="_top">
	<%END IF%>
	<div class="infoLink"><img src="link.png" border="0"></div>
	</a>
	</td>
	<td align="left" valign="top" class="setTwo1"><div class="listMed"><%response.write (Names("IndividualName"))%></div><%IF Names("Title") > "" Then%><div class="listSmall"><%response.write (Names("Title"))%></div><%END IF%></td>
	<td align="left" valign="top" class="setTwo2"><div class="listMed"><%response.write (Names("CompanyName"))%></div></td>
	<td align="center" valign="top" class="setTwo3"><div class="listMed"><%response.write (Names("Suite"))%></div></td>
	<td align="center" valign="top" class="setTwo4"><div class="listMed"><a href="indivanimatedmap.asp?ID=<%response.write(Names("IndID"))%>" target="_top"><img src="btn_map.jpg" border="0"></a></div></td>
	<td align="center" valign="top" class="setTwo5"><div class="listMed"></div></td>
</tr>

<% Names.movenext
loop %>
<%END IF%>

<%'Next, put in some empty rows to allow for additional scrolling
For count = 0 to 0%>
		<tr>
			<td colspan="5" height="45"><div class="listMed"></div></td>
		</tr>
<% next%>

</table>


<% conn.close
set conn = nothing
%>

</body>
</html>
