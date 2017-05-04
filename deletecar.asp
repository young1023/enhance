<!--#include file="database.inc"-->
 <%

 if session("carno")=""then
   response.redirect"searchresult.asp"
 else
   strsql="delete * from AccList where A_id= "& session("carno") & " "
		conn.execute strsql
		set strsql=nothing
		session("carno")=""
		response.redirect"enquiry.asp"
 end if
 
%>