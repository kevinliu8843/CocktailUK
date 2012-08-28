<%
Set FSO = Server.CreateObject("Scripting.FileSystemObject")
If NOT FSO.FileExists(Server.MapPath("/images/shop/products/"&Request("img"))) Then
	Set objGet= Server.CreateObject("MSXML2.ServerXMLHTTP")
	objGet.open "GET", "http://www.drinkstuff.com/productimg/" & Request("img"), False
	objGet.send ""
	call SaveBinaryData(Server.MapPath("/images/shop/products/"&Request("img")), objGet.ResponseBody)
End If
Set FSO = nothing
response.redirect("/images/shop/products/"&Request("img"))

Function SaveBinaryData(FileName, ByteArray)
  Const adTypeBinary = 1
  Const adSaveCreateOverWrite = 2
  
  'Create Stream object
  Dim BinaryStream
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To save binary data.
  BinaryStream.Type = adTypeBinary
  
  'Open the stream And write binary data To the object
  BinaryStream.Open
  BinaryStream.Write ByteArray
  
  'Save binary data To disk
  BinaryStream.SaveToFile FileName, adSaveCreateOverWrite
End Function
%>