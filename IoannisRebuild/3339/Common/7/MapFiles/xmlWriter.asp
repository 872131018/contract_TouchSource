<%@Language=Vbscript%>
<%option explicit%>
<%

  Dim byteCount, binRead, nonBinString, qFolderName, qFileName, fs, fo, fi

  qFolderName = Trim(Request("foldername"))
  qFileName = Trim(Request("filename"))
  
  byteCount = Request.TotalBytes
  binRead = Request.BinaryRead(byteCount)

  Function BinaryToString(binaryIn)
    Dim I, S, nextChar
    For I = 1 To LenB(binaryIn)
      nextChar = Chr(AscB(MidB(binaryIn, I, 1)))
      S = S & nextChar
    Next
    BinaryToString = S
  End Function

  nonBinString = BinaryToString(binRead)

  Set fs = Server.CreateObject("scripting.fileSystemObject")

  Set fo = fs.GetFolder(qFolderName)
  Set fi = fo.CreateTextFile(qFileName,true)

  fi.WriteLine(nonBinString)

  fi.Close
  Set fi = nothing
  Set fo = Nothing
  Set fs = nothing

  Response.Write("xmlWriter: "&qFileName&" done.")
%>
