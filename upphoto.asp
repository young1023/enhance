<!--#include file="database.inc"-->
<% 
id = session("id")
PhotoType = Request("PhotoType")

%>
<html>
<head>
<title>Shell Gas Maintenance Database</title>
<link rel="stylesheet" type="text/css" href="hse.css" />
<meta http-equiv="Refresh" content="2; url='upload_photo.asp?flag=<% =id %>'">
</head>

<body leftmargin="0" topmargin="0" >

     <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="10">
         <tr valign="top">
          <td height="100%" align="middle">
<%
 

'  Variables
'  *********
   Dim mySmartUpload
   Dim file
   Dim oConn
   Dim oRs
   Dim intCount
   Dim item
   Dim PhotoType
   intCount=0
        
'  Object creation
'  ***************
   Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

'  Upload
'  ******
   mySmartUpload.Upload


'  Open a recordset
'  ****************
   strSQL = "SELECT Order_id,name,path, PhotoType FROM Photo"
   'response.write strSQL 
   'response.end

   Set oRs = Server.CreateObject("ADODB.recordset")
   Set oRs.ActiveConnection = Conn
   oRs.Source = strSQL
   oRs.LockType = 3
   oRs.Open

'  Select each file
'  ****************
   For each file In mySmartUpload.Files
   '  Only if the file exist
   '  **********************
      If not file.IsMissing Then

      '  Add the current file in a DB field
      '  **********************************
         oRs.AddNew
		 oRs("Path") = file.FileName
		 oRs("Order_ID") = Session("ID")
		 oRs("PhotoType") = 1
         intCount = mySmartUpload.Save("Photo")
             oRs.Update
    
        intCount = intCount + 1
      End If
   Next

'  Display the number of files uploaded
'  ************************************
   Response.Write (intCount & " file(s) uploaded.<BR>")

   
'  Destruction
'  ***********
   oRs.Close
   Conn.Close
   Set oRs = Nothing 
   Set Conn = Nothing 

'-----------------------------------------------------------------------------
'
'      End of the main Content 
'
'-----------------------------------------------------------------------------
%>
</td>
              </tr>
                </table>
              </body>
              </html>