<!--#include file="database.inc"-->
<%
response.expires=0
 if  session("em")="" or session("id")="" then
        response.redirect "gologin.asp"
    end if
pageno=trim(request("pageno"))
 id=trim(request("chkid"))
 stracton=trim(request("action_button"))
if stracton="delete" then
        strsql="update member set group_id=0 where member.group_id in ("& id &") "
		'response.write strsql
		conn.execute strsql
		strsql="delete * from  user_group  where  group_id in("& id &") "
     	' response.write strsql
	    
		conn.execute strsql
		set acres=nothing
		
end if
 %>

<html>
<head>
<title>delete group</title>
<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type><LINK 
href="mailbox.files/imp.css" rel=stylesheet text="text/css">
<META content="MSHTML 5.00.2614.3500" name=GENERATOR>
<script language="JavaScript">
<!--
function dodelete()
{
	var i,j,k,strmsg;
	k=0;
	strmsg="";
	if (document.frmdelete.chkid!=null)
	{
		for(i=0;i<document.frmdelete.chkid.length;i++)
		{
			if(document.frmdelete.chkid[i].checked)
			k++;
		}
		if(i>0 && k==0)
		strmsg="You must select one group to delete!";
		if(i==0)
		{
			if (!document.frmdelete.chkid.checked)
			strmsg="You must select one group to delete!";
		}
	}
	if (i==0)
	  if (!(document.frmselect.chkid.checked))
	   strmsg="You must select one group to delete!";
	if (strmsg!="")
	alert(strmsg);
	else
	 if (confirm("Delete Group ,Yes Or No?"))
	{
		document.frmdelete.action_button.value="delete";
		document.frmdelete.submit();
	}
}
function doselectall()
{var i;
if (document.frmdelete.chkall.checked)
  for(i=0;i<document.frmdelete.chkid.length;i++)
     document.frmdelete.chkid[i].checked=true;
else
  for(i=0;i<document.frmdelete.chkid.length;i++)
     document.frmdelete.chkid[i].checked=false;
if (i==0)
	{if (document.frmdelete.chkall.checked)
	 document.frmdelete.chkid.checked=true;
	 else
	    document.frmdelete.chkid.checked=false;}
}

//-->
</script>

</HEAD>
<BODY aLink=#0000ff bgColor=#FFFFFF link=#0000ff text=#000000 vLink=#0000CC>
<p>&nbsp;</p>
<form name="frmdelete" method="post" action="delete_group.asp">
  <table width="100%" border="0" cellpadding="0">
    <tr bordercolor="#000000" bgcolor="#336699"> 
      <td colspan="3" height="20"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif">Group 
        Name</font></td>
    </tr>
    <% strsql="SELECT group_id,group_name from user_group order by group_name"
				set acres=nothing
				set acres=createobject("adodb.recordset")
				acres.cursortype=1
				acres.locktype=3
				acres.open strsql,conn
				'response.write strsql
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
    <tr bordercolor="#000000" bgcolor="#FFFFCC"> 
      <td width="4%" ><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
        <input type="checkbox" name="chkid" value="<%=acres("group_id")%>">
        </font></td>
      <td width="96%" ><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"><%=acres("group_name")%></font></td>
    </tr>
    <%acres.movenext  
	j=j+1 
	loop %>
    <tr bordercolor="#000000" bgcolor="#006699"> 
      <td colspan="3"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
        <input type="checkbox" name="chkall" value="checkbox" onClick="javascript:doselectall();">
        <font color="#FFFFFF"> Select All </font></font></td>
    </tr>
    <tr bordercolor="#000000" bgcolor="#FFFFCC"> 
      <td colspan="3"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
        <input type="button" name="cmdedit" value="Delete" onClick="javascript:dodelete();">
        <input type="hidden" name="action_button" value="">
        <input type="reset" name="cmdcancle" value="Reset">
        </font></td>
    </tr>
    <font face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
    <%	call showpageno(pageno) %>
    </font> 
     <tr bordercolor="#000000" bgcolor="#FFFFCC"> 
      <td colspan="3">
        <div align="right"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
          <%call showpageno(pageno)%>
          </font></div>
      </td>
    </tr> 
	<% else%>
  
    <tr bordercolor="#000000" bgcolor="#FFFFCC">
      <td colspan="3"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000">No 
        Record</font></td>
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
		     response.write "<a href='delete_group.asp?pageno=1'> The First Page</a>&nbsp;&nbsp;"
			response.write "<a href='delete_group.asp?pageno="&(pageno-19-(pageno  mod 20) )&"'>Previous 20</a>&nbsp;&nbsp;"
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
				response.write "<a href='delete_group.asp?Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;"
			end if	
		next
		if (pageno\20)<(lastpage\20) then
		        response.write "<a href='delete_group.asp?Pageno="&(pageno+21-(pageno mod 20)) &"'>Next 20</a>&nbsp;&nbsp;"
			   response.write "<a href='delete_group.asp?Pageno="& (lastpage) &"'> The Last Page</a>&nbsp;&nbsp;"
		end if
		
 end if
end function
%>
