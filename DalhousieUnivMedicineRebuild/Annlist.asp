<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>Directory Solutions - Amenity List</title>
	<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="No-Cache">
	<META HTTP-EQUIV="Expires" CONTENT="-1">
	<link rel="stylesheet" href="style_ds.css">
</head>

<% todaydate = FormatDateTime(Date(), 0)%>

<% timeminus = DateAdd("h", 2, Time())%>
<% chkStart = FormatDateTime(timeminus, 0)%>

<% timeplus = DateAdd("h", -.25, Time())%>
<% chkEnd = FormatDateTime(timeplus, 0)%>


	<%
	Session.timeout = 15
	Set conn = server.createObject("ADODB.Connection")
	conn.open "DalhousieUnivMedicine"
	Set Session("MyDB_conn")=conn
	%>

		
	<%
	StrSQLQuery = "SELECT * FROM Announcements WHERE #"&TodayDate&"# >= DisplayStart  ORDER BY StartDate ASC, StartTime ASC"
	Set qMeetings = Server.CreateObject("ADODB.recordset")
	qMeetings.Open strSQLQuery, conn, 3, 3
	%>	
	
<body background="bg_list.jpg" bgproperties="fixed" leftmargin="0" topmargin="0">
<div class="listXS">&nbsp;</div>

<table border="0" cellpadding="0" cellspacing="0" width="101%">

<% 'Put in this section a 'no matches' message if there are no listings with the selected letter
IF qMeetings.recordcount = 0 THEN%>
			<tr>
				<TD align="center" valign="top" class="setFive0"></td>
				<TD align="left" valign="middle" colspan="5" class="setFive6"><div class="listMed">There are no listings</div></TD>
			</TR>
<%
ELSE
%>

<%'If there are records to display, then start displaying them. %>
<% do until qMeetings.EOF%>


<TR>
	<TD align="center" valign="top" class="setFive0">
<!-- Check to see if there is a jpg or pdf, set href to point to those pages, otherwise, it goes to the detail page via the link icon. -->
	<%IF qMeetings("AnnounceImageName") > "" Then%>
		<% DIM AnnImageName %>
		<% AnnImageName = (qMeetings("AnnounceImageName"))%>
		<% DIM AnnImageLast3 %>
		<% AnnImageLast3 = right(AnnImageName, 3) %>
		<%IF AnnImageLast3 = "pdf" Then%>
			<a href="AnnPDFDisplay.asp?ID=<%response.write(qMeetings("ID"))%>" target="_top">
		<%ELSEIF AnnImageLast3 = "jpg" Then%>
			<a href="AnnJPGDisplay.asp?ID=<%response.write(qMeetings("ID"))%>" target="_top">
		<%END IF%>
	<%ELSE%>
		<a href="Anndetail.asp?ID=<%response.write(qMeetings("ID"))%>" target="_top">
	<%END IF%>
	<img src="link.png" border="0">
	</a>
	</td>
	<TD align="left" valign="top" class="setFive1"><div class="listMed"><%response.write (qMeetings("Name"))%></div></td>
	<TD align="left" valign="top" class="setFive2"><div class="listMed"><%response.write (qMeetings("Location"))%></div></td>
	<TD align="center" valign="top" class="setFive3"><div class="listSmall"><%response.write (qMeetings("StartDate"))%><% IF qMeetings("StartDate") <> qMeetings("EndDate") THEN %> -<br><%response.write (qMeetings("EndDate"))%><%END IF%></div></td>
	<TD align="center" valign="top" class="setFive4"><div class="listSmall"><%response.write (mid(qMeetings("StartTime"),1,4))%>	
		<%response.write (mid(qMeetings("StartTime"),8,9))%>
		-
		<%response.write (mid(qMeetings("EndTime"),1,4))%>
		<%response.write (mid(qMeetings("EndTime"),8,9))%></div></TD>
	<TD align="center" valign="top" class="setFive5"><div class="listMed"></div></TD>
</tr>

<% qMeetings.movenext
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
