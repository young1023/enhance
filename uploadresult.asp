<!--#include file="database.inc"-->
<%
if session(("SECURITY_LEVEL"))<>"1" then
				response.redirect "default.asp"
end if
   Dim mySmartUpload
   Dim ObjFile
   Dim intCount
   Dim StrMsg
   dim Strname(2)
   strname(1)=""
   strname(2)=""
   strprodid=session("product_id")
   Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")
   mySmartUpload.MaxFileSize =204800
   mySmartUpload.Upload
	
	i=0
    For each file In mySmartUpload.Files
		i=i+1
      If not file.IsMissing Then
	  	intCount = mySmartUpload.Save(server.mappath("uploadpic"))
		 strname(i)=file.fileName
	  end if
	next
	if strprodid<>"" then
		Set cmd = createobject("ADODB.Command")
		for i=1 to 2
			if strname(i)<>"" then
				if i=1 then
					strpic="update product set pic2path='uploadpic/"& strname(i) & "' where id="& strprodid
				else
					strpic="update product set picpath='uploadpic/"& strname(i) & "' where id="& strprodid
				end if
				cmd.CommandText = strpic
				cmd.ActiveConnection =Conn
				cmd.Execute
			end if
		next
		session("product_id")=""
	end if
	StrMsg="Picture Upload Success!"
%>
<html>		
<head>
<title>Struct</title>
<meta http-equiv="refresh" content="3;URL=DatAdmin.asp">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body  leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="images/demo/bdg1.jpg">
<table border="0" cellspacing="0" cellpadding="0" width="100%" height="455">

  <tr>
    <td width="644" valign="top">
      <div align="center"><b><%= strmsg %> </b></div>
    </td>
  </tr>
</table>
</body>
</html>
