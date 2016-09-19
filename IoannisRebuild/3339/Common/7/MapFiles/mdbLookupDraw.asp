<%@Language=Vbscript%>
<%Option Explicit%>
<%
  Dim qCompID, qBldg, qSuite, qMDBName
  
  qCompID = Trim(Request("sendCompID"))
  qBldg = Trim(Request("sendBldg"))
  qSuite = Trim(Request("sendSuite"))
  qMDBName = Trim(Request("mdbName"))

  Dim myConnection
  Set myConnection = Server.CreateObject("ADODB.Connection")
    myConnection.ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};" & "DBQ=" & Server.MapPath(qMDBName)
    myConnection.Open
    
  Dim compInfoSQL, myRS, CompName, BldgName, SuiteName, mainMessage
  
  Function CharCheck(nameIn)
    nameIn = Replace(nameIn, "&", "spchr038")
    nameIn = Replace(nameIn, "è", "spchr232")
    nameIn = Replace(nameIn, "é", "spchr233")
    nameIn = Replace(nameIn, "É", "spchr201")
    nameIn = Replace(nameIn, "â", "spchr226")
    nameIn = Replace(nameIn, "ä", "spchr228")
    nameIn = Replace(nameIn, "à", "spchr224")
    nameIn = Replace(nameIn, "á", "spchr225")
    nameIn = Replace(nameIn, "ì", "spchr236")
    nameIn = Replace(nameIn, "í", "spchr237")
    nameIn = Replace(nameIn, "î", "spchr238")
    nameIn = Replace(nameIn, "Ö", "spchr214")
    nameIn = Replace(nameIn, "ö", "spchr246")
    nameIn = Replace(nameIn, "ò", "spchr242")
    nameIn = Replace(nameIn, "ó", "spchr243")
    nameIn = Replace(nameIn, "ô", "spchr244")
    nameIn = Replace(nameIn, "ù", "spchr249")
    nameIn = Replace(nameIn, "ú", "spchr250")
    nameIn = Replace(nameIn, "û", "spchr251")
    nameIn = Replace(nameIn, "ñ", "spchr241")
    nameIn = Replace(nameIn, "Ñ", "spchr209")
    nameIn = Replace(nameIn, "Ç", "spchr199")
    nameIn = Replace(nameIn, "ç", "spchr231")
    CompName = nameIn
  End Function

  If Not qCompID = "empty" Then
    compInfoSQL = "SELECT * FROM Companies WHERE CompanyID = " & qCompID
  ElseIf Not qBldg = "empty" And Not qSuite = "empty" Then
    compInfoSQL = "SELECT * FROM Companies WHERE Building = '" & qBldg & "' AND Suite LIKE '%" & qSuite & "%'"
  Else
    compInfoSQL = "SELECT * FROM Companies WHERE CompanyName = 'null'"   'prevents asp failure
  End If

  Set myRS = Server.CreateObject("ADODB.Recordset")
    myRS.Open compInfoSQL, myConnection, 1, 1

  If Not myRS.EOF Then
    CompName = myRS("CompanyName")
      CharCheck(CompName)
    BldgName = myRS("building")
    SuiteName = myRS("Suite")
  
    Dim RecCount
    RecCount = myRS.RecordCount
    If RecCount > 1 Then      'if there is more than one company in this bldg/suite location
      CompName = CompName & "#"
    End If
  
    mainMessage = "returnBuilding=" & BldgName & "&returnSuite=" & SuiteName & "&returnName=" & CompName & "&RecCount=" & RecCount
  Else
    mainMessage = "returnMsg=No records found"
  End If
  
  myRS.Close
  Set myRS = Nothing

  myConnection.Close
  Set myConnection = Nothing

  Response.Write(mainMessage)
%>
