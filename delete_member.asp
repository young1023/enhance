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
  strsql="delete * from member  where id in("& id &")"
'  response.write strsql
  conn.execute strsql  
 end if
%>

<html>
<head>
<title>Delete Member</title>
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
		 strmsg=("You must select  memeber to delete !\n");
	if (i==0)
	      if (!(document.frmselect.chkid.checked))
		     strmsg=("You must select  memeber to delete !\n");
   if (strmsg!="")
      alert(strmsg);
   else
       if (confirm("Delete Member ,Yes Or No?"))
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
<form name="frmselect" method="post" action="delete_member.asp">
  <table width="100%" border="0" cellpadding="0">
    <tr bgcolor="#336699" bordercolor="#CCCCCC"> 
      <td width="7%"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif"></font></td>
      <td width="22%"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif">Email</font></td>
      <td width="27%"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif">Country</font></td>
      <td width="44%"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif">Type</font></td>
    </tr>
    <% strsql="select member.id,member.email,member.country,member.businesstype from member where member.level<>1 order by member.email"
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
				   i=(pageno-1)*20
				   acres.move i
				 end if
	if not acres.eof then
	  	do while not acres.eof
		   if j>19 then exit do%>
    <tr bgcolor="#FFFFCC" bordercolor="#CCCCCC"> 
      <td width="7%"> <font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
        <input type="checkbox" name="chkid" value="<%=acres("id")%>">
        </font></td>
      <td width="22%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"><%=acres("email")%></font></td>
      <td width="27%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"><%=acres("country")%></font></td>
      <td width="44%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"><%=acres("businesstype")%></font></td>
    </tr>
    <%	acres.movenext
    j=j+1
	loop%>
    <tr bgcolor="#006699" bordercolor="#CCCCCC"> 
      <td colspan="4"> <font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
        <font color="#FFFFFF"> 
        <input type="checkbox" name="chkall" value="checkbox" onclick="javascript:doselectall();">
        Select All </font></font></td>
    </tr>
    <font face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
    <%call showpageno(pageno)%>
    </font> 
    <tr bgcolor="#FFFFCC" bordercolor="#CCCCCC"> 
      <td colspan="4" height="18"> 
        <div align="left"><font color="#000000"> 
          <input type="button" name="cmddelet" value="Delete" onClick="javascript:dodelete();">
          <font face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
          <input type="hidden" name="cmddelete" value="">
          </font></font><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
          </font></div>
      </td>
    <tr bgcolor="#FFFFCC" bordercolor="#CCCCCC"> 
      <td colspan="5"> 
        <div align="right"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
          <%call showpageno(pageno)
		%>
          </font><font color="#000000"> </font> </div>
      </td>
    </tr>
    <%else%>
    <tr bgcolor="#FFFFCC" bordercolor="#CCCCCC"> 
      <td colspan="4"><font color="#000000">No Record</font></td>
    </tr>
    <%end if%>
  </table>
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
		     response.write "<a href='delete_member.asp?pageno=1'> The First Page</a>&nbsp;&nbsp;"
			response.write "<a href='delete_member.asp?pageno="&(pageno-19-(pageno  mod 20) )&"'>Previous 20</a>&nbsp;&nbsp;"
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
				response.write "<a href='delete_member.asp?Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;"
			end if	
		next
		if (pageno\20)<(lastpage\20) then
		        response.write "<a href='delete_member.asp?Pageno="&(pageno+21-(pageno mod 20)) &"'>Next 20</a>&nbsp;&nbsp;"
			   response.write "<a href='delete_member.asp?Pageno="& (lastpage) &"'> The Last Page</a>&nbsp;&nbsp;"
		end if
		
 end if
end function
%>
