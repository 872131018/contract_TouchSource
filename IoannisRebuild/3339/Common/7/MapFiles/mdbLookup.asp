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
