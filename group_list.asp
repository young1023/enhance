<!--#include file="database.inc"-->
<%
response.expires=0
 if  session("em")="" or session("id")="" then
        response.redirect "gologin.asp"
    end if
pageno=trim(request("pageno"))
 groupname=trim(request("chkid"))
 stracton=trim(request("action_button"))
strgroup=replace(trim(request("txtgroupname")),"'","''")
strRemember=replace(trim(request("cmdremeber")),"'","''")

if stracton="add" then
        strgroup=replace(strgroup,"'","''")
	    strsql="update user_group set group_name= '"& strgroup & "' where group_name='"& strRemember &"' "
		'response.write strsql
		conn.execute strsql
		set acres=nothing
		
end if
 %>

<html>
<head>
<title>Group List</title>
<META content="text/html; charset=big5" http-equiv=Content-Type><LINK 
href="mailbox.files/imp.css" rel=stylesheet text="text/css">
<META content="MSHTML 5.00.2614.3500" name=GENERATOR>
<script language="JavaScript">
<!--
function doedit()
{
	var i,j,k,strmsg;
	k=0;
	strmsg="";
	if (document.frmedit.chkid!=null)
	{
		for(i=0;i<document.frmedit.chkid.length;i++)
		{
			if(document.frmedit.chkid[i].checked)
			k++;
		}
		if(i>0 && k==0)
		strmsg="You have to select one group";
		if (k>1)
		strmsg="You can only select one group"
		if(i==0)
		{
			if (!document.frmedit.chkid.checked)
			strmsg="You have to select one group to edit";
		}
	}
	if (strmsg!="")
	alert(strmsg);
	else
	{
		document.frmedit.submit();
	}

}

function doupdate()
{
if((document.frmadd.txtgroupname.value)=="")
		alert("Please Input group Name!");
	else
		{
		document.frmadd.action_button.value="add";
		document.frmadd.submit();
		}
}

//-->
</script>

</HEAD>
<BODY aLink=#0000ff bgColor=#FFFFFF link=#0000ff text=#000000 vLink=#0000aa>
<%if groupname<>"" then%>
<form name="frmadd" method="post" action="group_list.asp">
  <table width="100%" border="0" cellpadding="0">
     
	
	 <tr bgcolor="#FFFFCC" bordercolor="#CCCCCC"> 
      <td width="20%" bgcolor="#006699"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif">Group 
        Name:</font></td>
      <td width="34%" bgcolor="#FFFFFF"> <font face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
        <input type="text" name="txtgroupname" size="25" maxlength="50" value="<%=groupname%>">
        </font></td>
      <td width="46%" bgcolor="#FFFFFF"> 
        <input type="button" name="cmdok" value="Ok" onclick="javascript:doupdate();">
		<input type="hidden" name="action_button" value="">
		<input type="hidden" name="cmdremeber" value="<%=groupname%>">
      </td>
    </tr>
  </table>
</form>
<%groupname=""
else%>
<p>&nbsp;</p>
<form name="frmedit" method="post" action="group_list.asp">
  <table width="100%" border="0" cellpadding="0">
    <tr bordercolor="#000000" bgcolor="#ffffcc">
      <td colspan="3"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
    <% strsql="select group_id,group_name from user_group order by group_name"
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
		call showpageno(pageno) %>
    </font></td>
    </tr>
	<tr bordercolor="#000000" bgcolor="#006699"> 
      <td colspan="3" height="20"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif">Group 
        Name</font></td>
    </tr>
    <%
	if not acres.eof then
	  	do while not acres.eof
		   if j>19 then exit do%>
    <tr bordercolor="#000000" bgcolor="#FFFFCC"> 
      <td width="4%" ><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
        <input type="checkbox" name="chkid" value="<%=acres("group_name")%>">
        </font></td>
      <td width="96%" ><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"><%=acres("group_name")%></font></td>
    </tr>
    <%acres.movenext  
	j=j+1 
	loop %>
   
    <tr bordercolor="#000000" bgcolor="#FFFFCC"> 
      <td colspan="3"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
        <input type="button" name="cmdedit" value="Edit" onclick="javascript:doedit();">
        <input type="reset" name="cmdcancle" value="Reset">
        </font></td>
    </tr>
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
<%
end if%>
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
		     response.write "<a href='group_list.asp?pageno=1'> The First Page</a>&nbsp;&nbsp;"
			response.write "<a href='group_list.asp?pageno="&(pageno-19-(pageno  mod 20) )&"'>Previous 20</a>&nbsp;&nbsp;"
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
				response.write "<a href='group_list.asp?Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;"
			end if	
		next
		if (pageno\20)<(lastpage\20) then
		        response.write "<a href='group_list.asp?Pageno="&(pageno+21-(pageno mod 20)) &"'>Next 20</a>&nbsp;&nbsp;"
			   response.write "<a href='group_list.asp?Pageno="& (lastpage) &"'> The Last Page</a>&nbsp;&nbsp;"
		end if
		
 end if
end function
%>
