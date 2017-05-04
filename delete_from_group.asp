<!--#include file="database.inc"-->
<%response.expires=0
 if  session("em")="" or session("id")="" then
        response.redirect "gologin.asp"
    end if
pageno=trim(request("pageno"))
group_id=trim(request("selgroup"))
addtogroup=trim(request("cmdaddtogroup"))
strgroup=replace(trim(request("selectgroup")),"'","''")
'response.write "group_name=" &strgroup
id=trim(request("chkid"))
'response.write group_id & addtogroup & id

if addtogroup="addto" then
  strsql="update member set group_id=0  where id in("& id &")"
'  response.write strsql
  conn.execute strsql  
 ' response.redirect "add_to_group_success.asp"

end if
%>

<html>
<head>
<title>Delete from group</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script language="javascript">
function doaddtogroup()
{  	var i,j,k,strmsg;
	k=0;
	strmsg="";
	if (document.frmselect.chkid!=null)
		{{for(i=0;i<document.frmselect.chkid.length;i++)
			  {if(document.frmselect.chkid[i].checked)
				k++;	
			  }
			if(i>0 && k==0)
			strmsg="You must select one memeber to be delete from group!\n";
		}
	   if(document.frmselect.selgroup.value=="-1" )
				strmsg="You must select group to delete!\n";

	   if (i==0)
	      if (!(document.frmselect.chkid.checked))
	            strmsg="You must select one memeber to be delete from group!\n";
	   if (strmsg!="")
		  alert(strmsg);
	   else
		   if (confirm("Delete Member From Group ,Yes Or No?"))
			{ document.frmselect.cmdaddtogroup.value="addto";
			document.frmselect.submit();}
	}
	else
	  alert("No Member  To Delete");
}
function doselectall()
{var i;
//alert(document.frmselect.chkid.length+1);
i=0;
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
	    document.frmselect.chkid.checked=false;
	}
}
 function doselect()
 {var i,strgroup1;

 for(i=0;i<frmselect.selgroup.length;i=i+1)
  if (document.frmselect.selgroup.options[i].selected)
     {strgroup1=frmselect.selgroup.options[i].text;
	 document.frmselect.selectgroup.value=strgroup1;
	// alert(strgroup1);
	document.frmselect.submit();
	}
 }
</script>
</head>

<body bgcolor="#FFFFFF" text="#000033" link="#0000CC" vlink="#0000CC" alink="#0000CC">
<form name="frmselect" method="post" action="delete_from_group.asp">
  <table width="100%" border="0" cellpadding="0">
    <tr bgcolor="#336699" bordercolor="#CCCCCC"> 
      <td width="3%"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif"></font></td>
      <td width="21%"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif">Email</font></td>
      <td width="23%"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif">Group</font></td>
      <td width="23%"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif">Country</font></td>
      <td width="30%"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif">Type</font></td>
    </tr>
    <% strsql="select member.id,member.email,member.country,member.businesstype,user_group.group_name" _
	           & " from user_group,member where member.group_id=user_group.group_id and group_name='"& strgroup &"' "
				set acres=nothing
			'	response.write strsql
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
      <td width="3%"> <font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
        <input type="checkbox" name="chkid" value="<%=acres("id")%>">
        </font></td>
      <td width="21%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"><%=acres("email")%></font></td>
      <td width="23%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"><%=acres("group_name")%></font></td>
      <td width="23%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"><%=acres("country")%></font></td>
      <td width="30%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"><%=acres("businesstype")%></font></td>
    </tr>
    <%	acres.movenext
    j=j+1
	loop%>
    <tr bgcolor="#006699" bordercolor="#CCCCCC"> 
      <td width="3%"> <font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
        <input type="checkbox" name="chkall" value="checkbox" onclick="javascript:doselectall();">
        </font></td>
      <td colspan="4"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#FFFFFF">Select 
        All </font></td>
    </tr>
    <font face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
    <%call showpageno(pageno)%>
    </font> 
    <tr bgcolor="#FFFFCC" bordercolor="#CCCCCC"> 
      <td colspan="5" height="25"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
        <%call showpageno(pageno)
		%>
        </font></td>
    </tr>
    <%else%>
    <tr bgcolor="#FFFFCC" bordercolor="#CCCCCC"> 
      <td colspan="5"><font color="#000000">No Record</font></td>
    </tr>
    <%end if%>
    <tr bgcolor="#FFFFCC" bordercolor="#CCCCCC"> 
      <td colspan="3"> <font color="#000000" face="Verdana,Geneva,Arial,Helvetica,sans-serif">Please 
        Select Group</font> </td>
      <td width="23%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
        <select name="selgroup"  onChange="javascript:doselect();">
          <option selected><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000">Select 
          Group</font></option>
          <% strsql="select * from user_group order by group_name"
		     set acres=conn.execute(strsql)
			do while not acres.eof
			  strtemp=trim(acres("group_name")&"")
	           if strtemp=strgroup then		%>
          <option value="<%=acres("group_id")%>"  selected><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"><%=strtemp%></font></option>
          <%else%>
          <option value="<%=acres("group_id")%>"  ><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"><%=strtemp%></font></option>
          <%end  if
		     acres.movenext
		   loop  %>
        </select>
        <input type="hidden" name="selectgroup" value="">
        </font></td>
      <td width="30%"> 
        <div align="center"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#000000"> 
          <input type="button" name="Button" value="Delete From Group" onClick="javascript:doaddtogroup();">
          <input type="hidden" name="cmdaddtogroup" value="">
          </font></div>
      </td>
    </tr>
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
		     response.write "<a href='delete_from_group.asp?pageno=1'> The First Page</a>&nbsp;&nbsp;"
			response.write "<a href='delete_from_group.asp?pageno="&(pageno-19-(pageno  mod 20) )&"'>Previous 20</a>&nbsp;&nbsp;"
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
				response.write "<a href='delete_from_group.asp?Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;"
			end if	
		next
		if (pageno\20)<(lastpage\20) then
		        response.write "<a href='delete_from_group.asp?Pageno="&(pageno+21-(pageno mod 20)) &"'>Next 20</a>&nbsp;&nbsp;"
			   response.write "<a href='delete_from_group.asp?Pageno="& (lastpage) &"'> The Last Page</a>&nbsp;&nbsp;"
		end if
		
 end if
end function
%>

