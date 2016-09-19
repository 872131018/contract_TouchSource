<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>Directory Solutions - Company List</title>
	<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="No-Cache">
	<META HTTP-EQUIV="Expires" CONTENT="-1">
	<link rel="stylesheet" href="style_ds.css">
	<link rel="stylesheet" type="text/css" href="scriptFiles/lightbox/lightbox.css" media="screen" />
	<link rel="stylesheet" type="text/css" href="scriptFiles/infoboxStyles.css" media="screen" />
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
	StrSQLQuery = "SELECT * FROM Companies"
	Set CompCount = Server.CreateObject("ADODB.recordset")
	CompCount.Open strSQLQuery, conn, 3, 3
	%>

	<%
	StrSQLQuery = "SELECT * FROM Companies WHERE CompanyName LIKE '"&request.querystring("letter")&"%'	 AND Description = 'R' ORDER BY CompanyName ASC"
	Set Companies = Server.CreateObject("ADODB.recordset")
	Companies.Open strSQLQuery, conn, 3, 3
	%>

<table border="0" cellpadding="0" cellspacing="0" width="101%">

<% 'Put in this section a 'no matches' message if there are no listings with the selected letter
IF Companies.recordcount = 0 THEN%>
		<tr>
				<TD align="center" valign="top" class="setOne0"></td><TD align="left" colspan="5" class="setOne6"><div class="listLarge"><% IF (CompCount.recordcount = 0) AND (request.querystring("letter") = "") THEN%>There are no listings at this time<%ELSE%>There are no listings that start with "<%response.write (request.querystring("letter"))%>"<%END IF%></div></TD>
   		</TR>
<%
ELSE
%>

<%'If there are records to display, then start displaying them. %>
<% do until Companies.EOF%>

<TR>
	<TD align="center" valign="middle" class="setOne0">
		<div>
			<a class="wayfinder" href="#" target="_top" data-company="<%response.write(Companies("CompanyID"))%>">
				<img src="btn_map.jpg" border="0">
			</a>
		</div>
	</td>
	<TD align="left" valign="middle" class="setOne1"><div class="listLarge"><%response.write (Companies("CompanyName"))%></div></td>
	<TD align="left" valign="middle" class="setOne2"><div class="listLarge"><%response.write (Companies("Suite"))%></div></td>
	<TD align="left" valign="middle" class="setOne3"><div class="listLarge"><%response.write (Companies("Building"))%></div></td>
	<TD align="center" valign="middle" class="setOne4"><div class="listLarge"></div></TD>
	<TD align="center" valign="middle" class="setOne5"><div class="listLarge"></div></TD>
</tr>

<% Companies.movenext
loop %>
<%END IF%>

<%'Next, put in some empty rows to allow for additional scrolling
For count = 0 to 15%>
<TR>
	<TD colspan="6" height="45"><div class="listLarge"></div></TD>
</tr>
<% next%>

</TABLE>

<% conn.close
set conn = nothing
%>
<% '/*
'*Move the scripts to the bottom of the body so they are non blocking
'*/ %>
<script language="javascript" src="scriptFiles/jquery-2.1.4.js"></script>
<script language="javascript" src="scriptFiles/jquery.cookie.js"></script>
<script language="javascript" src="scriptFiles/fabric.js"></script>
<script language="javascript" src="scriptFiles/lightbox/lightbox.js"></script>
<script>
// Wait until the DOM has loaded before querying the document
$(document).ready(function()
{
	/*
	* Restore the state of the window for better experience
	*/
	setTimeout(function()
	{
		$("html, body").scrollTop($.cookie("currentScroll"));
	}, 200)
	/*
	* Each button image  will make ajax request for map
	*/
	$(".wayfinder").on('click', function(event)
	{
		/*
		* Save the window scroll state when the user finds the map
		*/
		$.cookie('currentScroll', window.scrollY);
		/*
		* Load the animated map and launch it in a modal
		*/
		$.get('animatedmap.asp?Companies='+$(this).data("company"), function(data)
		{
			modal.open({content: data});
		});
	});
});
</script>
</body>
</html>
