<%@Language=Vbscript%>
<%Option Explicit%>
<%
  Dim qCompID, qIndID, qMDBName
  qCompID =Trim(Request("compID"))
  qIndID =Trim(Request("indID"))
  qMDBName =Trim(Request("mdbName"))

  Dim myConnection
  Set myConnection=Server.CreateObject("ADODB.Connection")
    myConnection.ConnectionString="DRIVER={Microsoft Access Driver (*.mdb)};" & "DBQ=" & Server.MapPath(qMDBName)
    myConnection.Open
    
  Dim compInfoSQL, myRS, CompName, iCompID, iIndivName, mainMessage
  
  Function CharCheck(nameIn)
    nameIn = Replace(nameIn, "&", "spchr038")
    nameIn = Replace(nameIn, "�", "spchr232")
    nameIn = Replace(nameIn, "�", "spchr233")
    nameIn = Replace(nameIn, "�", "spchr201")
    nameIn = Replace(nameIn, "�", "spchr226")
    nameIn = Replace(nameIn, "�", "spchr228")
    nameIn = Replace(nameIn, "�", "spchr224")
    nameIn = Replace(nameIn, "�", "spchr225")
    nameIn = Replace(nameIn, "�", "spchr236")
    nameIn = Replace(nameIn, "�", "spchr237")
    nameIn = Replace(nameIn, "�", "spchr238")
    nameIn = Replace(nameIn, "�", "spchr214")
    nameIn = Replace(nameIn, "�", "spchr246")
    nameIn = Replace(nameIn, "�", "spchr242")
    nameIn = Replace(nameIn, "�", "spchr243")
    nameIn = Replace(nameIn, "�", "spchr244")
    nameIn = Replace(nameIn, "�", "spchr249")
    nameIn = Replace(nameIn, "�", "spchr250")
    nameIn = Replace(nameIn, "�", "spchr251")
    nameIn = Replace(nameIn, "�", "spchr241")
    nameIn = Replace(nameIn, "�", "spchr209")
    nameIn = Replace(nameIn, "�", "spchr199")
    nameIn = Replace(nameIn, "�", "spchr231")
    CompName = nameIn
  End Function

  If qCompID <> "undefined" Then
    compInfoSQL = "SELECT * FROM Companies WHERE CompanyID = " & qCompID
    
    Set myRS=Server.CreateObject("ADODB.Recordset")
      myRS.Open compInfoSQL, myConnection
      
    CompName = myRS("CompanyName")
    
    CharCheck(CompName)
    
    mainMessage="Building=" & myRS("building") & "&Suite=" & myRS("Suite") & "&CompanyName=" & CompName
      
  Else
    If qIndID <> "undefined" Then
    
      compInfoSQL = "SELECT * FROM individuals WHERE IndId = " & qIndID
      
      Set myRS=Server.CreateObject("ADODB.Recordset")
       myRS.Open compInfoSQL, myConnection
      
      iIndivName = myRS("IndividualName")
      iCompID = myRS("CompanyID")
      
      myRS.Close
      Set myRS=Nothing
      
      compInfoSQL = "SELECT * FROM Companies WHERE CompanyID = " & iCompID
    
      Set myRS=Server.CreateObject("ADODB.Recordset")
        myRS.Open compInfoSQL, myConnection
      
      mainMessage = "Building=" & myRS("building") & "&Suite=" & myRS("Suite") & "&IndividualName=" & iIndivName
    End If
  End If

myRS.Close
Set myRS=Nothing
myConnection.Close
Set myConnection=Nothing

'Response.ContentType="application/x-www-form-urlencoded"
Response.Write(mainMessage)
%>
