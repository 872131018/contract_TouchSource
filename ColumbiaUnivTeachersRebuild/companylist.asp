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
  '/*
  '*Create a connection and a session for accessing the db
  '*/
  Set connection = Server.CreateObject("ADODB.Connection")
  connection.Open "ColumbiaUnivTeachers"
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
    <TR>
    	<TD align="left" valign="top" class="setOne1"><div class="listLarge"><%response.write (Companies("CompanyName"))%></div></td>
    	<TD align="center" valign="top" class="setOne2"><div class="listLarge"><%response.write (Companies("Building"))%></div></td>
    	<TD align="center" valign="top" class="setOne3"><div class="listLarge"><%response.write (Companies("Suite"))%></div></td>
    	<TD align="center" valign="top" class="setOne4"><div class="listLarge"><% IF Companies("Building") > "" THEN %><% IF Companies("Building") <> "N/A" THEN %>
        <!--<a href="companyanimatedmap.asp?ID=<%response.write(Companies("CompanyID"))%>&BLDG=<%response.write(Companies("Building"))%>&ROOM=<%response.write(Companies("Suite"))%>&FL=<%response.write(Companies("Floor"))%>" target="_top">-->
        <a class="wayfinder" href="#" target="_top" data-company="<%response.write(Companies("CompanyID"))%>">
          <img src="btn_map.jpg" border="0">
        </a>
        <%END IF%><%END IF%></div></TD>
    	<TD align="center" valign="top" class="setOne5"><div class="listLarge"></div></TD>
    </tr>
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
