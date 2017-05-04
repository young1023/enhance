<!--#include file="database.inc"-->
<%
response.expires=0
 if  session("em")="" or session("id")="" then
        response.redirect "gologin.asp"
 end if
pageno=trim(request("pageno")) %>

<html>
<head>
<title>Group List</title>
<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type><LINK 
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
//-->
</script>

</HEAD>
<BODY aLink=#0000ff bgColor=#FFFFFF link=#0000ff text=#000000 vLink=#0000aa>
<form name="frmedit" method="post" action="updata.asp">
  <table width="100%" border="0" cellpadding="0">
    <tr bordercolor="#000000" bgcolor="#006699"> 
      <td width="3%" height="2" ><font color="#FFFFFF"></font></td>
      <td height="2" width="10%"><font color="#FFFFFF">E_mail</font></td>
      <td height="2" width="9%"><font color="#FFFFFF">Sex</font></td>
      <td width="19%" height="2"><font color="#FFFFFF">Country</font></td>
      <td width="27%" height="2"><font color="#FFFFFF">Business Type</font></td>
      <td width="32%" height="2"><font color="#FFFFFF">Join Date</font></td>
    </tr>
    <% strsql="select * from member order by email"
				set acres=nothing
				set acres=createobject("adodb.recordset")
				acres.cursortype=1
				acres.locktype=3
				acres.open strsql,conn
			'	response.write strsql
				recount=acres.recordcount
			    response.write "Total Member"& recount
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
      <td width="3%" ><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
        <input type="checkbox" name="chkid" value="<%=acres("id")%>">
        </font></td>
      <td width="10%" ><font color="#000000"><%=acres("email")%></font></td>
      <td width="9%" ><font color="#000000"><%=acres("sex")%></font></td>
      <td width="19%" ><font color="#000000"><%=acres("country")%> </font></td>
      <td width="27%" ><font color="#000000"><%=acres("businesstype")%> </font></td>
      <td width="32%" ><%=acres("Create_date")%></td>
    </tr>
    <%acres.movenext  
	j=j+1 
	loop %>
    <font face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
    <%	call showpageno(pageno) %>
    </font> 
    <tr bordercolor="#000000" bgcolor="#FFFFCC"> 
      <td colspan="7"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
        <input type="button" name="cmdedit" value="Edit" onclick="javascript:doedit();">
        <input type="reset" name="cmdcancle" value="Reset">
        </font></td>
    </tr>
    <tr bordercolor="#000000" bgcolor="#FFFFCC"> 
      <td colspan="7"> 
        <div align="right"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
          <%call showpageno(pageno)%>
          </font></div>
      </td>
    </tr>
    <% else%>
    <tr bordercolor="#000000" bgcolor="#FFFFCC"> 
      <td colspan="7"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000">No 
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
		     response.write "<a href='edit_member.asp?pageno=1'> The First Page</a>&nbsp;&nbsp;"
			response.write "<a href='edit_member.asp?pageno="&(pageno-19-(pageno  mod 20) )&"'>Previous 20</a>&nbsp;&nbsp;"
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
				response.write "<a href='edit_member.asp?Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;"
			end if	
		next
		if (pageno\20)<(lastpage\20) then
		        response.write "<a href='edit_member.asp?Pageno="&(pageno+21-(pageno mod 20)) &"'>Next 20</a>&nbsp;&nbsp;"
			   response.write "<a href='edit_member.asp?Pageno="& (lastpage) &"'> The Last Page</a>&nbsp;&nbsp;"
		end if
		
 end if
end function
%>