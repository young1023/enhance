<!--#include file="database.inc"-->
<%response.expires=0
 if  session("em")="" or session("id")="" then
        response.redirect "gologin.asp"
    end if
pageno=trim(request("pageno"))
delete=trim(request("cmddelete"))
id=trim(request("chkid"))

'response.write delete & id

if delete="delete" then
  strsql="delete * from memery_email  where email_id in("& id &")"
'  response.write strsql
  conn.execute strsql  
 end if
%>

<html>
<head>
<title>Delete E_mail</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script language="javascript">
function dodelete()
{  	var i,j,k,strmsg;
	k=0;
	strmsg="";
	if (document.frmselect.chkid!=null)
	{for(i=0;i<document.frmselect.chkid.length;i++)
		  if(document.frmselect.chkid[i].checked)
			k++;	
	}
	if(i>0 && k==0)
		 strmsg=("You must select  information to delete !\n");
	if (i==0)
	      if (!(document.frmselect.chkid.checked))
		     strmsg=("You must select  information to delete !\n");
   if (strmsg!="")
      alert(strmsg);
   else
       if (confirm("Delete  Selected information ,Yes Or No?"))
	   { document.frmselect.cmddelete.value="delete";
		document.frmselect.submit();}
}
function doselectall()
{var i;
if (document.frmselect.chkall.checked)
  for(i=0;i<document.frmselect.chkid.length;i++)
     document.frmselect.chkid[i].checked=true;
else
  for(i=0;i<document.frmselect.chkid.length;i++)
     document.frmselect.chkid[i].checked=false;
if (i==0)
	{if (document.frmselect.chkall.checked)
	 document.frmselect.chkid.checked=true;
	 else
	    document.frmselect.chkid.checked=false;}
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#0000CC" vlink="#0000CC" alink="#0000CC">
<form name="frmselect" method="post" action="delete_email.asp">
  <table width="100%" border="0" cellpadding="0">
    <tr bgcolor="#336699" bordercolor="#CCCCCC"> 
      <td width="7%"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif"></font></td>
      <td width="22%"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif">To</font></td>
      <td width="27%"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif">Subject</font></td>
      <td width="27%"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif">No</font></td>
      <td><font color="#FFFFFF">Date</font></td>
    </tr>
    <% strsql="select  * from memery_email order by Send_date Desc"
				set acres=nothing
				'response.write strsql
				set acres=createobject("adodb.recordset")
				acres.cursortype=1
				acres.locktype=3
				acres.open strsql,conn
				
				recount=acres.recordcount
			
				if pageno="" then
				   pageno=1
			    end if
				 if pageno>1 then
				   i=(pageno-1)*10
				   acres.move i
				 end if
	if not acres.eof then
	  	do while not acres.eof
		   if j>9 then exit do%>
    <tr bgcolor="#FFFFCC" bordercolor="#CCCCCC"> 
      <td width="7%"> <font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
        <input type="checkbox" name="chkid" value="<%=acres("email_id")%>">
        </font></td>
      <td width="22%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"><%=acres("conto")%></font></td>
      <td width="27%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"><a href="send1.asp?radpop=<%=acres("confield")%>&who=<%=acres("conto")%>&txtsubject=<%=acres("email_subject")%>&textarea=<%=acres("email_body")%>"><%=acres("email_subject")%></a></font></td>
      <td width="27%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"><%=acres("send_number")%></font></td>
      <td><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"><%=acres("send_date")%></font></td>
    </tr>
    <%	acres.movenext
    j=j+1
	loop%>
    <tr bgcolor="#006699" bordercolor="#CCCCCC"> 
      <td colspan="5"> <font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
        <font color="#FFFFFF"> 
        <input type="checkbox" name="chkall" value="checkbox" onclick="javascript:doselectall();">
        Select All </font></font></td>
    </tr>
    <font face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
    <%call showpageno(pageno)%>
    </font> 
    <tr bgcolor="#FFFFCC" bordercolor="#CCCCCC"> 
      <td colspan="5" height="18"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
        <%call showpageno(pageno)
		%>
        </font></td>
    <tr bgcolor="#FFFFCC" bordercolor="#CCCCCC"> 
      <td colspan="6"> <font color="#000000"> 
        <input type="button" name="cmddelet" value="Delete" onclick="javascript:dodelete();">
        <font face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
        <input type="hidden" name="cmddelete" value="">
        </font></font> </td>
    </tr>
    <%else%>
    <tr bgcolor="#FFFFCC" bordercolor="#CCCCCC"> 
      <td colspan="5"><font color="#000000">No Record</font></td>
    </tr>
    <%end if%>
  </table>
</form>
<p>&nbsp;</p></body>
</html>
<% 
function showpageno(pageno)
	if recount>10 then
		lastpage=recount\10
		if recount mod 10 >0 then
			lastpage=lastpage+1
		end if
		if pageno>10 then
		     response.write "<a href='delete_email.asp?pageno=1'> The First Page</a>&nbsp;&nbsp;"
			response.write "<a href='delete_email.asp?pageno="&(pageno-9-(pageno  mod 10) )&"'>Previous 10</a>&nbsp;&nbsp;"
		end if
		strtemp=pageno
		if (pageno Mod 10 )=0 then
		   strtemp=strtemp-10
		end if
	 for i=(strtemp-(strtemp mod 10)+1) to (strtemp+10-(strtemp mod 10))
	         if lastpage<i then  exit for			 
            if i- pageno=0 then
				response.write cstr(i)&"&nbsp;&nbsp;"
			else
				response.write "<a href='delete_email.asp?Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;"
			end if	
		next
		if (pageno\10)<(lastpage\10) then
		        response.write "<a href='delete_email.asp?Pageno="&(pageno+11-(pageno mod 10)) &"'>Next 10</a>&nbsp;&nbsp;"
			   response.write "<a href='delete_email.asp?Pageno="& (lastpage) &"'> The Last Page</a>&nbsp;&nbsp;"
		end if
		
 end if
end function
%>