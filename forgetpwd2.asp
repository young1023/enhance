<!--#include file="database.inc"-->
<%
email=trim(request("email"))
if email="" then
  response.redirect "forgetpwd.asp"
end if
strpwd=""
straction=trim(request("action_button"))
if straction="getpwd" then
  response.expires=0
  answer=replace(trim(request("answer")),"'","''")
  strsql="select mypassword from member where email='"& email &"' and answer='"& answer &"' "
  'response.write strsql
  set acre=conn.execute(strsql)
  if not acre.eof then	
    strpwd=replace(trim(acre("mypassword")),"'","''")
    'response.write strpwd
  else
	strpwd="Sorry,the answer is wrong !<br><br>For further assistance,send us mail <a href='mailto:info@jaguarpen.com'>info@jaguarpen.com</a>"
  end if
  set acre=nothing
  set strsql=nothing
else
  sql="select question from member where email='"& email &"' "
  set w=conn.execute(sql)
  if not w.eof then
    question=replace(trim(w("question")),"'","''")
    'response.write question
    'response.end
	w.close
  else
    response.redirect "forgetpwd.asp"
  end if

end if

%>
<html>
<head>
<title>Forget password</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--
function dochkval()
{
	var errmsg="";
	if(document.frmpwd.answer.value=="")
		errmsg=errmsg+"Plesase Input Answer!\n";
	if (errmsg!="")
		alert(errmsg);
	else{
		document.frmpwd.action_button.value="getpwd";
		document.frmpwd.submit();
	}
	
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="images/demo/bdg1.jpg">
<table border="0" cellspacing="0" cellpadding="0" width="100%" height="455">
  <tr align="center"> 
    <td colspan="2" height="125"> 
      <!--#include file="top.inc" --></td>
  </tr>
  <tr> 
    <td valign="top" width="145" align="center" rowspan="2"> 
      <!--#include file="left.inc" --></td>
    <td height="146" width="644" valign="top"><!--#include file="sbar.inc" --></td>
  </tr>
  <tr>
    <td width="644" valign="top">
      <div align="center">
	  <% If strpwd="" then %>
	    <form name="frmpwd" method="post" action="forgetpwd2.asp">
          <table width="60%" border="1" cellspacing="1" cellpadding="0">
            <tr> 
              <td><b><font face="Arial" size="2" color="#EFFBBF"><strong>Question 
                :</strong></font></b></td>
              <td><b><%=question%>
                <input type=hidden name=email value=<%=email%>>
                </b></td>
            </tr>
            <tr> 
              <td><b>Answer to question :</b></td>
              <td><b> 
                <input type="text" name="answer" size="30" maxlength="50">
                </b></td>
            </tr>
            <tr> 
              <td colspan="2"> 
                <div align="center"><b>&nbsp;</b><b> 
                  <input type="button" name="Submit" value="Get Password" onclick="dochkval();">
                  <input type="reset" value="Reset Form">
                  <input type="hidden" name="action_button" value="">
                  </b></div>
              </td>
            </tr>
          </table>
        </form>
		<% Else %>
      <font size="2" face="Arial, Helvetica, sans-serif" color="#FFFFFF"> 
       </font><font size="4" face="Arial, Helvetica, sans-serif" color="#ffFFcc">
        Your Password is <%=strpwd %></font><font size="2" face="Arial, Helvetica, sans-serif" color="#FFFFFF"> 
  </font></div>
		  
		<%end if%>
    </td>
  </tr>
</table>
<div align="center"><font size="2" color="#CCCCCC">Copyright &copy; 2001 Winson 
  Development Company All right reserved.</font> </div>
</body>
</html>
