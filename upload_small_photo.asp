<!--#include file="database.inc"-->
<%
session("id") = request("flag")
'id = session("id")
%>
<html>
<head>
<title>Upload Photo</title>
<link rel="stylesheet" type="text/css" href="images/basic.css" />
</head>
<body leftmargin="0" topmargin="0">
      <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="2" class="Normal">
        <tr>
          <td align="middle">
請上戴細相片<br>
<FORM METHOD="POST" ACTION="upsmallphoto.asp" ENCTYPE="multipart/form-data">
   <input type="hidden" name="flag" size="23" value="<% = id %>" ><br><br>
   <INPUT TYPE="FILE" NAME="FILE1" SIZE="30"><p>
   <INPUT TYPE="submit" VALUE="      上戴    ">
</p>
</FORM>
</td>
  </tr>
 
<%
   Photo_Sql = " Select * from Photo where PhotoType = 2 and  Order_id = "&Session("id")
   Set oPRs = Conn.Execute(Photo_Sql)
      If Not oPRs.eof then
  do while not oPRs.eof
%>
<tr>
<td align="middle">
<a href="Photo/<% = oPRs("Path") %>" target="_blank"><img src="Photo/<% = oPRs("Path") %>" width="200" height="150"></a>
</td>
   </tr>
<%                         
                               oPRs.Movenext
							Loop
						End If
					%>
<tr><td align=center>
   <INPUT TYPE="BUTTON" VALUE="      Close Window" onClick="window.close()"></td></tr>
     </table>
 </body>
    </html>