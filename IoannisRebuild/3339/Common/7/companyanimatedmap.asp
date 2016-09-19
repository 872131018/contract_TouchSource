<!DOCTYPE html>
<html>
<head>
	<title>Directory Solutions - Company Detail Page</title>
	<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="No-Cache">
	<META HTTP-EQUIV="Expires" CONTENT="-1">
	<link rel="stylesheet" href="style_map.css">
</head>


<body background="bg_map.jpg" bgproperties="fixed" leftmargin="0" topmargin="0" class="overflowHidden">


	<%
	Session.timeout = 15
	Set conn = server.createObject("ADODB.Connection")
	conn.open "Update"
	Set Session("MyDB_conn")=conn
	%>
	
	<%
	StrSQLQuery = "SELECT *	FROM Companies WHERE CompanyID = "&request.querystring("ID")&""
	Set CompID = Server.CreateObject("ADODB.recordset")
	CompID.Open strSQLQuery, conn, 3, 3
	%>
	
	<%
	StrSQLQuery = "SELECT *	FROM List12 WHERE (Text7 LIKE '"&request.querystring("CODE")&"%')"
	Set List12ID = Server.CreateObject("ADODB.recordset")
	List12ID.Open strSQLQuery, conn, 3, 3
	%>


<DIV>
	<embed src="MapApp.swf?Companies=<%response.write(CompID("CompanyID"))%>"
	  quality="high" bgcolor="d4cfc9"
	  width="1040" height="750"
	  name="MapApp"
	  align="middle"
	  allowScriptAccess="sameDomain"
	  allowFullScreen="false"
	  type="application/x-shockwave-flash"
	  pluginspage="http://www.adobe.com/go/getflashplayer" />
</DIV>

<div class="mainTitle">
<%response.write (CompID("CompanyName"))%> - <%response.write (CompID("Suite"))%>
</div>
<div class="mainDirections">
<%response.write (List12ID("Memo1"))%> 
</div>	
	<% conn.close
	set conn = nothing
	%>	

	
</body>
</html>
