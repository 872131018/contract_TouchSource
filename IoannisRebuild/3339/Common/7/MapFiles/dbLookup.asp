<%@Language=Vbscript%>
<%Option Explicit
  On Error Resume Next
  
  Dim mainMessage
  mainMessage = "Begin=Now"
  Dim errorCount
  errorCount = 0
  
  Function CheckError()
    errorCount = errorCount + 1
    If Err.Number <> 0 Then
      'Error.Clear
      mainMessage = mainMessage & "&ErrorCheck" & errorCount & "=" & Err.Description
    Else
      mainMessage = mainMessage & "&ErrorCheck" & errorCount & "=NoErrorDetected"
    End If
  End Function
  
  Dim qDataSource, qSearchTable, qIDField, qSearchID, qBldgField, qSuiteField
  Dim qDvar1, qDvar2, qDvar3, qDvar4, qDvar5, qDvar6, qDvar7, qDvar8, qDvar9, qDvar10

  qDataSource = Trim(Request("dataSource"))
  qSearchTable = Trim(Request("searchTable"))
  
  qIDField =  Trim(Request("idField"))
  qSearchID = Trim(Request("searchId"))
  
  qBldgField = Trim(Request("bldgField"))
  qSuiteField = Trim(Request("suiteField"))
  
  qDvar1 = Trim(Request("dvar1"))
  qDvar2 = Trim(Request("dvar2"))
  qDvar3 = Trim(Request("dvar3"))
  qDvar4 = Trim(Request("dvar4"))
  qDvar5 = Trim(Request("dvar5"))
  qDvar6 = Trim(Request("dvar6"))
  qDvar7 = Trim(Request("dvar7"))
  qDvar8 = Trim(Request("dvar8"))
  qDvar9 = Trim(Request("dvar9"))
  qDvar10 = Trim(Request("dvar10"))
  
  CheckError()
  
  Dim conn
  Set conn = Server.CreateObject("ADODB.Connection")
  
  conn.ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};" & "DBQ=" & Server.MapPath(qDataSource)
    'conn.ConnectionString = "Provider=MSDASQL;" & "DSN=" & qDataSource
    'conn.ConnectionString = "DSN=" & qDataSource
    
  conn.mode = 3
  conn.Open
  
  CheckError()
  
  Dim searchSQL, rs1, rs2, retSuite, retBldg, compID
  Dim displayVar1, displayVar2, displayVar3, displayVar4, displayVar5, displayVar6, displayVar7, displayVar8, displayVar9, displayVar10
  
  '''''''replace special characters with Flash charCodes
  
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
    CharCheck = nameIn
  End Function
  
  '''''''retrieve data for map: suite and building
  
  searchSQL = "SELECT * FROM " & qSearchTable & " WHERE " & qIDField & " = " & qSearchID
  Set rs1 = Server.CreateObject("ADODB.Recordset")
  rs1.Open searchSQL, conn
    
  If rs1.EOF Then
    mainMessage = mainMessage & "&MapSearchIDNotFound=" & qSearchID
  Else
    If qBldgField <> "empty" Then
      mainMessage = mainMessage & "&Building=" & rs1(qBldgField)
    Else
      mainMessage = mainMessage & "&Building=empty"
    End If
    If qSuiteField <> "empty" Then
      mainMessage = mainMessage & "&Suite=" & rs1(qSuiteField)
    Else
      mainMessage = mainMessage & "&Suite=empty"
    End If
  End If
  
  rs1.Close
  Set rs1 = Nothing
  
  '''''''retrieve data for infobox
  
  searchSQL = "SELECT * FROM " & qSearchTable & " WHERE " & qIDField & " = " & qSearchID
  Set rs2=Server.CreateObject("ADODB.Recordset")
  rs2.Open searchSQL, conn
  
  If rs2.EOF Then
    If mainMessage <> "" Then
      mainMessage = mainMessage & "&"
    End If
    mainMessage = mainMessage & "InfoboxSearchIDNotFound=" & qSearchID
  Else
  
    If qDvar1 <> "empty" Then
      displayVar1 = rs2(qDvar1)
      displayVar1 = CharCheck(displayVar1)
      If mainMessage <> "" Then
        mainMessage = mainMessage & "&"
      End If
      mainMessage = mainMessage & "displayVar1=" & displayVar1
    End If
    
    If qDvar2 <> "empty" Then
      displayVar2 = rs2(qDvar2)
      displayVar2 = CharCheck(displayVar2)
      If mainMessage <> "" Then
        mainMessage = mainMessage & "&"
      End If
      mainMessage = mainMessage & "displayVar2=" & displayVar2
    End If
    
    If qDvar3 <> "empty" Then
      displayVar3 = rs2(qDvar3)
      displayVar3 = CharCheck(displayVar3)
      If mainMessage <> "" Then
        mainMessage = mainMessage & "&"
      End If
      mainMessage = mainMessage & "displayVar3=" & displayVar3
    End If
    
    If qDvar4 <> "empty" Then
      displayVar4 = rs2(qDvar4)
      displayVar4 = CharCheck(displayVar4)
      If mainMessage <> "" Then
        mainMessage = mainMessage & "&"
      End If
      mainMessage = mainMessage & "displayVar4=" & displayVar4
    End If
    
    If qDvar5 <> "empty" Then
      displayVar5 = rs2(qDvar5)
      displayVar5 = CharCheck(displayVar5)
      If mainMessage <> "" Then
        mainMessage = mainMessage & "&"
      End If
      mainMessage = mainMessage & "displayVar5=" & displayVar5
    End If
    
    If qDvar6 <> "empty" Then
      displayVar6 = rs2(qDvar6)
      displayVar6 = CharCheck(displayVar6)
      If mainMessage <> "" Then
        mainMessage = mainMessage & "&"
      End If
      mainMessage = mainMessage & "displayVar6=" & displayVar6
    End If
    
    If qDvar7 <> "empty" Then
      displayVar7 = rs2(qDvar7)
      displayVar7 = CharCheck(displayVar7)
      If mainMessage <> "" Then
        mainMessage = mainMessage & "&"
      End If
      mainMessage = mainMessage & "displayVar7=" & displayVar7
    End If
    
    If qDvar8 <> "empty" Then
      displayVar8 = rs2(qDvar8)
      displayVar8 = CharCheck(displayVar8)
      If mainMessage <> "" Then
        mainMessage = mainMessage & "&"
      End If
      mainMessage = mainMessage & "displayVar8=" & displayVar8
    End If
    
    If qDvar9 <> "empty" Then
      displayVar9 = rs2(qDvar9)
      displayVar9 = CharCheck(displayVar9)
      If mainMessage <> "" Then
        mainMessage = mainMessage & "&"
      End If
      mainMessage = mainMessage & "displayVar9=" & displayVar9
    End If
    
    If qDvar10 <> "empty" Then
      displayVar10 = rs2(qDvar10)
      displayVar10 = CharCheck(displayVar10)
      If mainMessage <> "" Then
        mainMessage = mainMessage & "&"
      End If
      mainMessage = mainMessage & "displayVar10=" & displayVar10
    End If
    
  End If
  
  rs2.Close
  Set rs2 = Nothing
  
  conn.Close
  Set conn = Nothing
  
  CheckError()
  
  Response.write(mainMessage)
%>
