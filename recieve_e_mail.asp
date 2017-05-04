<!--#include file="database.inc"-->
<%
    if  session("em")="" or session("id")="" then
        response.redirect "gologin.asp"
    end if
response.expires=0
chkid=trim(request("chkid"))
delet=trim(request("cmddeletehid"))
pageno=trim(request("pageno"))
strsql="select * from admin"
set acres=conn.execute(strsql)
set pop3=server.createobject("Jmail.pop3")
name=trim(acres("name")&"")
password=trim(acres("password")&"")
pop= trim(acres("pop3")&"")
pop3.connect name,password,pop
if delet="delet" then
   	dim aryid
	aryid=split(chkid,",")
	for i=0 to ubound(aryid)
  		pop3.deletesinglemessage aryid(i)
  	next
	response.redirect "delete_email_success.asp"
end if
%>
<html>
<head>
<title>Recieve E_mail</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script language="javascript">
<!--
function doselectall()
{var i;
if (document.frmrecieve.chkall.checked)
  for(i=0;i<document.frmrecieve.chkid.length;i++)
     document.frmrecieve.chkid[i].checked=true;
else
  for(i=0;i<document.frmrecieve.chkid.length;i++)
     document.frmrecieve.chkid[i].checked=false;
if (i==0)
	{if (document.frmrecieve.chkall.checked)
	 document.frmrecieve.chkid.checked=true;
	 else
	    document.frmrecieve.chkid.checked=false;}
}

function dodelete()
{  	var i,j,k,strmsg;
	k=0;
	strmsg="";
	if (document.frmrecieve.chkid!=null)
	{for(i=0;i<document.frmrecieve.chkid.length;i++)
		  if(document.frmrecieve.chkid[i].checked)
			k++;	
	}
	if(i>0 && k==0)
		 strmsg=("You must select  E_mail to delete !\n");
	if (i==0)
	      if (!(document.frmrecieve.chkid.checked))
		     strmsg=("You must select  E_mail to delete !\n");
   if (strmsg!="")
      alert(strmsg);
   else
       if (confirm("Delete  Selected E_mail ,Yes Or No?"))
	   { document.frmrecieve.cmddeletehid.value="delet";
	   
		document.frmrecieve.submit();}
}
-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#0000CC" vlink="#0000CC" alink="#0000CC">
<form name="frmrecieve" method="post" action="recieve_e_mail.asp">
  <%
strsql="select * from admin"
set acres=conn.execute(strsql)
set pop3=server.createobject("Jmail.pop3")
name=trim(acres("name")&"")
password=trim(acres("password")&"")
pop= trim(acres("pop3")&"")
pop3.connect  name,password,pop

 if pop3.count > 0 then
    Set msg = pop3.Messages.item(1) 
   ReTo = ""
   Set Recipients = msg.Recipients
%>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="30">
    <tr> 
      <td colspan="5" height="19"><font color="#FF0000"> 
        <%response.write("you have "& pop3.count &" E_mail in you mailbox!<br><br>")%>
        </font></td>
    </tr>
    <tr> 
      <td colspan="5" height="14"> 
        <div align="left"><font color="#FF0000"> 
          <%if pageno="" then
	              pageno=1
			end if
			counti=1
			 if pageno>1 then
					 counti=(pageno-1)*20
			 end if
				recount=pop3.count
		  call showpageno(pageno)%>
          </font> </div>
      </td>
    </tr>
    <tr bgcolor="#006699"> 
      <td width="85" height="21">&nbsp;</td>
      <td width="203" height="21"> 
        <div align="left"><font color="#FFFFFF">From</font></div>
      </td>
      <td width="251" height="21"><font color="#FFFFFF">Date</font></td>
      <td width="57" height="21"> 
        <div align="left"><font color="#FFFFFF">Subject</font></div>
      </td>
      <td width="183" height="21"> </td>
    </tr>
    <%
  if pop3.count>0 then
   For i =counti To pop3.Count
     '  response.write "i="&i&"count="&pop3.count
       if j>19 then exit for
	   Set msg=pop3.messages.item(i)
   %>
    <tr bgcolor="#FFFFCC"> 
      <td width="85" height="21"> 
        <input type="checkbox" name="chkid" value="<%=i%>">
      </td>
      <td width="203" height="21"> 
        <div align="left"><font color="#000000"><%= msg.FromName %></font></div>
      </td>
      <td width="251" height="21"><font color="#000000"><%= msg.date %></font></td>
      <td colspan="2" height="21"> 
        <div align="left"><font color="#000000"><a href="emailbody.asp?sendsubject=<%=server.urlencode( msg.Subject) %>&sendfrom=<%=server.urlencode(msg.FromName) %>&sendDate=<%=server.urlencode( msg.date) %>&hiddenbody=<%=server.urlencode(msg.Body)%>"><%= msg.Subject %></a> 
          </font></div>
      </td>
    </tr>
    <%
	j=j+1
	Next %>
    <tr bgcolor="#006699"> 
      <td width="85" height="21"> <font color="#FFFFFF"> 
        <input type="checkbox" name="chkall" value="checkbox" onclick="javascript:doselectall();">
        </font></td>
      <td width="203" height="21"> 
        <div align="left"><font color="#000000"><font color="#000000"><font color="#FFFFFF">Select 
          All </font></font></font></div>
      </td>
      <td width="251" height="21"><font color="#FFFFFF"></font></td>
      <td colspan="2" height="21"> 
        <div align="left"><font color="#000000"><font color="#000000"><font color="#FFFFFF"></font></font></font></div>
      </td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td colspan="5" height="21"> 
        <div align="right"><font color="#FF0000"> 
          <%call showpageno(pageno)%>
          </font></div>
      </td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td width="85" height="27"> 
        <input type="button" name="cmddelete" value="Delete" onclick="javascript:dodelete();">
        <input type="hidden" name="cmddeletehid" value="">
      </td>
      <td width="203" height="27">&nbsp;</td>
      <td width="251" height="27">&nbsp;</td>
      <td colspan="2" height="27">&nbsp;</td>
    </tr>
    <%else%>
    <tr bgcolor="#FFFFCC"> 
      <td colspan="5" height="18"> No Record </td>
    </tr>
    <%end if%>
  </table> 
   <%else
   response.write "Your Mailbox is empty!"%> 
     
  <%end if  %> 
</form> 
<p>&nbsp;</p></body> 
</html> 
 
<%  
function showpageno(pageno) 
	if recount>20 then 
		lastpage=recount\20 
		if recount mod 20 >0 then 
			lastpage=lastpage+1 
		end if 
		if pageno>10 then 
		     response.write "<a href='recieve_e_mail.asp?pageno=1'> The First Page</a>&nbsp;&nbsp;" 
			response.write "<a href='recieve_e_mail.asp?pageno="&(pageno-19-(pageno  mod 20) )&"'>Previous 20</a>&nbsp;&nbsp;" 
		end if 
		strtemp=pageno 
		if (pageno Mod 20 )=0 then 
		   strtemp=strtemp-20 
		end if 
	 for i=(strtemp-(strtemp mod 20)+1) to (strtemp+20-(strtemp mod 20)) 
	         if lastpage<i then  exit for			  
            if i- pageno=0 then 
				response.write cstr(i)&"&nbsp;&nbsp;" 
			else 
				response.write "<a href='recieve_e_mail.asp?Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;" 
			end if	 
		next 
		if (pageno\20)<(lastpage\20) then 
		        response.write "<a href='recieve_e_mail.asp?Pageno="&(pageno+21-(pageno mod 20)) &"'>Next 20</a>&nbsp;&nbsp;" 
			   response.write "<a href='recieve_e_mail.asp?Pageno="& (lastpage) &"'> The Last Page</a>&nbsp;&nbsp;" 
		end if 
		 
 end if 
end function 
%>