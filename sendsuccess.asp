<!--#include file="database.inc"-->
<%
 if  session("em")="" or session("id")="" then
        response.redirect "gologin.asp"
    end if
strto=trim(request("txtto"))
email_array=split(strto,",")
strsubject=trim(request("txtsubject"))
strcontent=trim(request("txtcontent"))
strsql="select * from admin"
	set acres=conn.execute(strsql)
	Set jmail = Server.CreateObject("JMail.Message")
	for i=0 to ubound(email_array)
	  jmail.AddRecipient chr(34)&email_array(i)&chr(34)
	next
	jmail.From =chr(34)&trim(acres("mailbox")&"")&chr(34)
	jmail.Subject=chr(34)&trim(strsubject)&chr(34)
	strbody=strcontent
	jmail.Body=chr(34)&trim(strbody)&chr(34)
	strsend= trim(acres("smtp")&"")
	issuccess=false
	if not jmail.Send(strsend) then
	  response.write "Your SMTP Server set up error,Please set again !<br><br>"
	  response.write "Or can not connect SMTP Server!"
	else
		issuccess=true
	 response.write "Email sent succesfully!" 
   end if
%>
	
<html>
<head>
<title>Send Email</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#0000CC" vlink="#0000CC" alink="#0000CC">
<%if issuccess then%>
<table width="288">
  <tr> 
    <td width="149" bgcolor="#006699"><font color="#FFFFFF">Subject</font></td>
    <td width="127" bgcolor="#FFFFCC"><font color="#000000"><%= strSubject %></font></td>
  </tr>
  <tr> 
    <td width="149" bgcolor="#006699"><font color="#FFFFFF">From</font></td>
    <td width="127" bgcolor="#FFFFCC"><font color="#000000"><%= trim(acres("smtp") &"") %></font></td>
  </tr>
  <tr> 
    <td width="149" bgcolor="#006699"><font color="#FFFFFF">Body</font></td>
    <td width="127" bgcolor="#FFFFCC"> <font color="#000000"><%=strcontent %> 
      </font></td>
  </tr>
</table>
<%end if%>
</body>
</html>
