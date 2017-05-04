<%
if session(("SECURITY_LEVEL"))<>"1" then
  response.redirect "default.asp"
end if
session("picflage")=""
session("product_id")=""
%>
<!--#include file ="included/conne.inc" -->
<%
response.expires=0
whatdo=trim(request.form("whatdo"))
Select case whatdo
 case "del"
   delid=split(trim(request.form("mid")),",")
   for i=0 to Ubound(delid)
     sql="delete id from nomination where id="&delid(i)&" and employeenum='"&session("u_enum")&"' "
     conn.execute(sql)
	 'response.write sql&"<br>"
   next
   mymessage="Delete the Record Successfully"
   whatgo="user_edit.asp"
 case "delmember"
  if session("shell_power")=8 or session("shell_power")=3 then
   delid=split(trim(request.form("mid")),",")
   for i=0 to Ubound(delid)
     sql="delete id from member where id="&delid(i)&" and flage<>8 "
     conn.execute(sql)
	 'response.write sql&"<br>"
   next
  end if
   mymessage="Delete the Member Successfully"
   whatgo="sa_member.asp"

'======================================================
' Delete News
'======================================================


 case "delnews"
    delid=split(trim(request.form("mid")),",")
   for i=0 to Ubound(delid)
     sql="delete id from message where id="&delid(i)
     conn.execute(sql)
	 'response.write sql&"<br>"
   next
   mymessage="The Newsletter(s) was/were deleted successfully"
   whatgo="sa_news.asp"

'======================================================
' Delete Tradeshow
'======================================================


 case "delshow"
   delid=split(trim(request.form("mid")),",")
   for i=0 to Ubound(delid)
     sql="delete id from tradeshow where id="&delid(i)
     conn.execute(sql)
	 'response.write sql&"<br>"
   next
   mymessage="The tradeshow(s) record was/were deleted successfully"
   whatgo="sa_tradeshow.asp"


 case "delnews"
 if session("shell_power")=3 or session("shell_power")=8 then
   delid=split(trim(request.form("mid")),",")
   for i=0 to Ubound(delid)
     sql="delete id from message where id="&delid(i)&" "
     conn.execute(sql)
     'response.write sql&"<br>"
   next
 end if
   mymessage="Delete News Successfully"
   whatgo="sa_news.asp"
 case "setmodel"
  if session("shell_power")=8 or session("shell_power")=3 then
   sql="Update modul set flage=0"
   conn.execute(sql)
   doing=split(trim(request.form("doing")),",")
   for i=0 to Ubound(doing)
     sql="Update modul set flage=1 where id="&doing(i)&" "
     conn.execute(sql)
   next
  end if
   mymessage="Set Model Active Successfully"
   whatgo="listbill.asp"
case ""
  conn.close
  set conn=nothing
  response.redirect "user.asp"
End select

conn.close
set conn=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Refresh" content="3; url='<%=whatgo%>'">
<title>Shell</title>
</head>
<body topmargin="0" marginwidth="0" marginheight="0" leftmargin="0" >
<br><br>
<table border=0 cellpadding=3 cellspacing=0 class=hardcolor width="90%" align=center>
  <tbody>
  <tr> 
    <td align=center bgcolor="#006699" height="23"><font color=white><%=mymessage%></font></td>
    </tr>
    </tbody>
  </table>
<br><br>
<div align="center">
<script language=JavaScript src="included/copyright.js"></script>
</div>
</body>
</html>