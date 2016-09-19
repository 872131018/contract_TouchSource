<!DOCTYPE HTML>
<html>
<head>
	<title>Directory Solutions - Amenity List</title>
	<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="No-Cache">
	<META HTTP-EQUIV="Expires" CONTENT="-1">
	<link rel="stylesheet" href="style_ds.css">
</head>

<body background="bg_list2.jpg" bgproperties="fixed" leftmargin="0" topmargin="0" class="listingsWidth">
<div class="pushdown"></div>

	<%
	Session.timeout = 15
	Set conn = server.createObject("ADODB.Connection")
	conn.open "100CorporatePointe"
	Set Session("MyDB_conn")=conn
	%>
	
	<%
	StrSQLQuery = "SELECT *	FROM Amenities ORDER BY Val (Amenities.Phone) DESC"
	Set qAmenities = Server.CreateObject("ADODB.recordset")
	qAmenities.Open strSQLQuery, conn, 3, 3
	%>


<table id="tbl_listing" border="0" cellpadding="0" cellspacing="0" class="listingsWidth">

<% 'Put in this section a 'no matches' message if there are no listings with the selected letter
IF qAmenities.recordcount = 0 THEN%>
			<tr>
				<td align="right" valign="top" class="setThree0"></td>
				<td align="left" valign="middle" colspan="5" class="setThree6"><div class="listMed">There are no listings</div></td>
			</tr>
<%
ELSE
%>

<%'If there are records to display, then start displaying them. %>
<% do until qAmenities.EOF%>

<tr>
	<td align="center" valign="top" class="setThree0">
	<!-- Check to see if there is a jpg or pdf, set href to point to those pages, otherwise, it goes to the detail page via the link icon. -->
	<%IF qAmenities("PhotoFile") > "" Then%>
		<% DIM AmenImageName %>
		<% AmenImageName = (qAmenities("PhotoFile"))%>
		<% DIM AmenImageLast3 %>
		<% AmenImageLast3 = right(AmenImageName, 3) %>
		<%IF AmenImageLast3 = "pdf" Then%>
			<a href="AmenPDFDisplay.asp?ID=<%response.write(qAmenities("AmenityID"))%>" target="_top">
		<%ELSEIF AmenImageLast3 = "jpg" Then%>
			<a href="AmenJPGDisplay.asp?ID=<%response.write(qAmenities("AmenityID"))%>" target="_top">
		<%END IF%>
	<%ELSE%>
		<a href="Amenitydetail.asp?ID=<%response.write(qAmenities("AmenityID"))%>" target="_top">
	<%END IF%>
	<div class="infoLink"><img src="link.png" border="0"></div>
	</a>
	</td>
	<td align="left" valign="top" class="setThree1"><div class="listMed"><%response.write (qAmenities("AmenityName"))%></div></td>
	<td align="left" valign="top" class="setThree2"><div class="listMed"><%response.write (qAmenities("Location"))%><%IF qAmenities("Description") > "" Then%></div><div class="listSmall"><%response.write (qAmenities("Description"))%><%END IF%></div></td>
	<td align="left" valign="top" class="setThree3"><div class="listMed"></div></td>
	<td align="center" valign="top" class="setThree4"><div class="listMed"></div></td>
	<td align="center" valign="top" class="setThree5"><div class="listMed"></div></td>
</tr>

<% qAmenities.movenext
loop %>
<%END IF%>

<%'Next, put in some empty rows to allow for additional scrolling
For count = 0 to 0%>
<tr>
	<td colspan="5" height="45"><div class="listLarge"></div></td>
</tr>
<% next%>

</table>


<% conn.close
set conn = nothing
%>

</body>
</html>
