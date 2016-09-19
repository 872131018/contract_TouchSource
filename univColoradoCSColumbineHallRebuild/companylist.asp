<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>Directory Solutions - Company List</title>
	<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="No-Cache">
	<META HTTP-EQUIV="Expires" CONTENT="-1">
	<link rel="stylesheet" type="text/css" href="style_ds.css"  media="screen">
</head>

<body leftmargin="0" topmargin="0">
	<%
  '/*
  '*Create a connection and a session for accessing the db
  '*/
  Set connection = Server.CreateObject("ADODB.Connection")
  connection.Open "UnivColoradoCSColumbineHall"
  Session.timeout = 15
  Set Session("MyDB_conn")=connection
	%>
	<%
  '/*
	'*Set query, create a record set, execute query, and count results
	'*@TODO: collapse ifelse to only contain logic for creating sql query to reduce code duplication
	'*@note: reassigning queryString
	'*/
	queryString = "SELECT * FROM Companies WHERE CompanyName LIKE '"&request.querystring("letter")&"%'	ORDER BY CompanyName ASC"
	Set Companies = Server.CreateObject("ADODB.recordset")
	Companies.Open queryString, connection, 3, 3
	%>

<!--magic here that I haven't really looked into -->
<table border="0" cellpadding="0" cellspacing="1" width="101%" class="OnePixelBorder">
<% 'Put in this section a 'no matches' message if there are no listings with the selected letter
IF Companies.recordcount = 0 THEN%>
		<tr class="bgRow1">
				<TD colspan="5" height="50"><div class="listLarge">&nbsp;&nbsp;There are no listings that start with this letter</div></TD>
   		</TR>
<%
i=1
ELSE
%>
<%'If there are records to display, then start displaying them.   Change row colors by incrementing the 'i' variable by one and resetting to zero if 2, so the only two options are 1 and 2%>
<% i=0
do until Companies.EOF%>
<a href="companydetail.asp?ID=<%response.write(Companies("CompanyID"))%>" target="_top">
<%i=i+1
if i=2 then
i=0
response.write "<TR class=bgRow1>"
ELSE
response.write "<TR class=bgRow2>"
end if
%>
<TD align="left" valign="middle" class="setOne1"><div class="listLarge"><%response.write (Companies("CompanyName"))%></div></td>
<TD align="center" valign="middle" class="setOne2"><div class="listLarge"><%response.write (Companies("Suite"))%></div></td>
<TD align="center" valign="middle" class="setOne3">
  <div class="listLarge">
  <a class="wayfinder" href="#" target="_top" data-company="<%response.write(Companies("CompanyID"))%>">
    <img src="btn_map.jpg" border="0">
  </a>
  </div>
</td>
<TD align="center" valign="middle" class="setOne4"><div class="listLarge"></div></TD>
<TD align="center" valign="middle" class="setOne5"><div class="listLarge"></div></TD>
</tr>
</a>
<% Companies.movenext
loop %>
<%END IF%>
<%'Next, put in some empty rows to allow for additional scrolling
For count = 0 to 15
i=i+1
if i=2 then
i=0
response.write "<TR class=bgRow1>"
else
response.write "<TR class=bgRow2>"
end if
%>
	<TD align="left" valign="middle" class="setOne1"><div class="listMed"></div></td>
	<TD align="left" valign="middle" class="setOne2"><div class="listMed"></div></td>
	<TD align="center" valign="middle" class="setOne3"><div class="listMed"></div></td>
	<TD align="center" valign="middle" class="setOne4"><div class="listMed"></div></td>
	<TD align="center" valign="middle" class="setOne5"><div class="listMed"></div></td>
</tr>
<% next%>
</TABLE>
<!-- end of the magic that I don't know about! -->
<%
'/*
'*Close all connections and release variables for garbage collection
'*/
connection.close
set connection = nothing
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
