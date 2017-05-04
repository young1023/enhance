<!--#include file="database.inc"-->
 <%
response.expires=0 
  if  session("SECURITY_LEVEL")<>1 then
         response.redirect "gologin.asp"
  end if

strsql="select * from admin"
set acres=conn.execute(strsql)

if not acres.eof then
strsmtp=trim(acres("smtp")&"")
strpop=trim(acres("pop3")&"")
strpassword=trim(acres("password")&"")
strname=trim(acres("name")&"")
strserveremail=trim(acres("mailbox")&"")
end if

strsuccess=trim(request("textarea"))
condition1=trim(request("radpop"))
who=trim(request("who"))

if condition1="" then
  condition1="alluser"
end if

select case  condition1
    case "group"
	   radpop="group"
	   emailcon=trim(request("txtgroupname"))
	case "country"
	   radpop="country"
	    emailcon=trim(request("txtcountry"))
	case "businesstype"
	   radpop="businesstype"
	    emailcon=trim(request("txttype"))
	case "comanysize"
	   radpop="comanysize"
	    emailcon=trim(request("txtcomsize"))
	case "alluser"
	   radpop="alluser"
	   emailcon=trim("AllMember")
	case "userinput"
	   radpop="userinput"
	   emailcon=trim(request("txtto"))	
end select

strSubject =trim(request("txtsubject"))
strcontent=trim(request("textarea"))
send=trim(request("send"))
'response.write radpop
issuccess=false
if send="send" then
      issuccess=true
      send=""
      select case radpop
		 case "group" 
			  strsqltemp="select group_id from user_group where  group_name='"& emailcon &"'"
			   'response.write strsqltemp
			  set acrestemp=conn.execute(strsqltemp)
			  if not acrestemp.eof then
				strsql="select email from member where subscribe=0 and email<>'' and group_id="& acrestemp.fields(0).value &""
			  else
			   response.redirect "nothisman.asp"
			  end if		 
		 case "country","businesstype","comanysize"
		     strsql="select email from member where  subscribe=0 and email<>'' and "& radpop &"='"& emailcon &"'"
	     case "alluser"
		     strsql="select email from member where subscribe=0 and email<>'' "
		 case "userinput"
		     strsql="select * from member where subscribe=0 and id=0"
	  end select	
'response.write strsql			
    set acres=conn.execute(strsql)	
	Set jmailmsg = Server.CreateObject("JMail.Message")
	email=""
	if acres.eof then	   
	    emailarray=split(emailcon,",")
		max=ubound(emailarray)
	    num=0
		for i=0 to max
		    email=trim(emailarray(i)&"")
	    	if instr(email,"@")=0  then
			  response.redirect "nothisman.asp"
			else
			  jmailmsg.AddRecipient chr(34)& email &chr(34)
			  num=num+1
			end if
	       ' response.write email
			jmailmsg.From =chr(34)& strserveremail &chr(34)
			jmailmsg.Subject=chr(34)&trim(strsubject)&chr(34)
			strbody=strcontent
			jmailmsg.Body=chr(34)&trim(strbody)&chr(34)
			
			jmailmsg.logging=true
			sendflag=true
			'response.write strsmtp
			if not jmailmsg.Send(strsmtp) then
		       sendflag=false
			end if 
			jmailmsg.clearrecipients
		next 	
	else
	   num=0
	   do while not acres.eof
		    email=trim(acres.fields(0).value&"")	
			'response.write email
		    jmailmsg.AddRecipient chr(34)& email &chr(34)
		  	jmailmsg.From =chr(34)& strserveremail &chr(34)
			jmailmsg.Subject=chr(34)&trim(strsubject)&chr(34)
			strbody=strcontent
			jmailmsg.Body=chr(34)&trim(strbody)&chr(34)
		
	    	jmailmsg.logging=true
			sendflag=true
			if not jmailmsg.Send(strsmtp) then
			  sendflag=false
			end if 
			num=num+1
	        acres.movenext
			jmailmsg.clearrecipients
	   loop 
       set jmailmsg=nothing	   
    end if
 
if not sendflag then
  response.write "Your SMTP Server set up error,Please set again !<br><br>"
  response.write "Or can not connect SMTP Server!"
else
  response.write "Email sent successfully!"
end if

strsql="insert into memery_email(confield,conto,email_subject,email_body,send_number)" _
			& " values('"& condition1 &"', '"& emailcon &"','"& strsubject &"','"& strcontent &"',"& num &")"
conn.execute strsql			
'response.write strsql
end if

%> 
 <html>
<head>
<title>Send E_mail</title>
<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type><LINK 
href="mailbox.files/imp.css" rel=stylesheet text="text/css">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
</head>

<body bgcolor="#003366" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFCC">
<%if issuccess then%>
<table width="288" border="1">
  <tr bgcolor="#003366" bordercolor="#FFFFFF"> 
    <td width="149"><font color="#FFFF99">Subject</font></td>
    <td width="127"><font color="#FFFF99" face="&lt;font face=&quot;Verdana,Geneva,Arial,Helvetica,sans-serif&quot; color=&quot;#FFFFFF&quot;&gt; "><%= strSubject %>&nbsp;&nbsp;</font></td>
  </tr>
  <tr bgcolor="#003366" bordercolor="#FFFFFF"> 
    <td width="149"><font color="#FFFF99">From</font></td>
    <td width="127"><font color="#FFFF99" face="&lt;font face=&quot;Verdana,Geneva,Arial,Helvetica,sans-serif&quot; color=&quot;#FFFFFF&quot;&gt; "><%= strserveremail %>&nbsp;&nbsp;</font></td>
  </tr>
  <tr bgcolor="#003366" bordercolor="#FFFFFF"> 
    <td width="149"valign="top"><font color="#FFFF99">Body</font></td>
    <td height="128" width="576"> <font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#FFFF99"> 
      <textarea disabled name="txtcontent" rows="6" cols="50"><%=strsuccess%></textarea>
      </font></td>
  </tr>
</table>
<%end if%> 
</body> 
</html>
