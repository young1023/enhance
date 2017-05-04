<!--#include file="database.inc"-->
<%
strorderpage=trim(request("strorder"))
if strorderpage="" then
  strorderpage=1
end if
pageno=trim(request("pageno"))
groupname=replace(trim(request("txtgroupname")),"'","''")
country=replace(trim(request("txtcountry")),"'","''")
name=replace(trim(request("txtname")),"'","''")
email=replace(trim(request("txtemail")),"'","''")
strtype=replace(trim(request("txttype")),"'","''")
companyname=replace(trim(request("txtcomname")),"'","''")
companysize=replace(trim(request("txtcomsize")),"'","''")
sex=trim(request("radmr"))
strsql="SELECT  member.*, user_group.*" _
  & "FROM user_group RIGHT JOIN member ON user_group.group_id = member.Group_Id" _
  & " WHERE ((((member.Group_Id)=user_group.group_id)) OR ((member.group_id=0)))"
'   response.write strsql
		 if groupname<>"" then
		    strsql=strsql & "  and user_group.group_name='"& groupname &" ' "
		end if
		if country<>"" then
		    strsql=strsql & "  and member.country='"& country &" ' "
		end if
		if name<>"" then
		    strsql=strsql & "  and member.loginname='"& name &" ' "
		end if
		if email<>"" then
		    strsql=strsql & "  and member.email='"& email &" ' "
		end if
		if strtype<>"" then
		    strsql=strsql & "  and member.busiesstype='"& strtype &" ' "
		end if
		if companyname<>"" then
		    strsql=strsql & "  and member.companyname='"& companyname &" ' "
		end if
		if companysize<>"" then
		    strsql=strsql & "  and member.companysize='"& companysize &" ' "
       end if
	   if sex="ms" then
	         strsql=strsql & "  and member.sex='ms' "
	   end if
	   if sex="mr" then
	        strsql=strsql & "  and member.sex='mr' "
	   end if
	  	  select case strorderpage
	    case 1
		   strorder="loginname"
		case 2
		  strorder="email"
        case 3
		   strorder="group_name"
		case 4
		  strorder="country"
		case 5
		    strorder="businesstype"
	end select 
	if strorderpage<>"" then
	    strsql=strsql & " order by "& strorder &""
	end if
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
				   i=(pageno-1)*10
				   acres.move i
				 end if%>

<html>
<head>
<title>GSelect Member</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script language="javascript">
function dosend()
{  	var i,j,k,strmsg;
	k=0;
	strmsg="";
	if (document.frmselect.chkid!=null)
	{for(i=0;i<document.frmselect.chkid.length;i++)
		  {if(document.frmselect.chkid[i].checked)
			k++;	
		  }
		if(i>0 && k==0)
		strmsg="You must select one memeber to send!\n";
	}
	if (i==0)
	  if (!(document.frmselect.chkid.checked))
	    strmsg="You must select one memeber to send!\n";
   if (strmsg!="")
	  alert(strmsg);
   else
		document.frmselect.submit();
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

<body bgcolor="#CCCCCC" text="#000000">
<form name="frmselect" method="post" action="send.asp">
  <table width="100%" border="0" cellpadding="0">
   <tr bgcolor="#FFFFCC" bordercolor="#CCCCCC"> 
      <td width="14%"><font color="#999999" face="Verdana,Geneva,Arial,Helvetica,sans-serif"></font></td>
      <td width="52%"><font color="#999999" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><a href="Select%20Member.asp?strorder=1">Name</a></font></td>
      <td width="52%"><font color="#999999" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><a href="ESelect%20Member.asp?strorder=2">Email</a></font></td>
      <td width="52%"><font color="#999999" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><a href="GSelect%20Member.asp?strorder=3">Group</a></font></td>
      <td width="52%"><font color="#999999" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><a href="CSelect%20Member.asp?strorder=4">Country</a></font></td>
      <td width="34%"><font color="#999999" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><a href="TSelect%20Member.asp?strorder=5">Type</a></font></td>
    </tr>
	<%	if not acres.eof then
	  	do while not acres.eof
		   if j>9 then exit do%>
    <tr bgcolor="#FFFFFF" bordercolor="#CCCCCC"> 
      <td width="14%"> <font face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
        <input type="checkbox" name="chkid" value="<%=acres("id")%>">
        </font></td>
      <td width="52%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif"><%=acres("loginname")%></font></td>
      <td width="52%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif"><%=acres("email")%></font></td>
      <td width="52%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif"><%=acres("group_name")%></font></td>
      <td width="52%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif"><%=acres("country")%></font></td>
      <td width="34%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif"><%=acres("businesstype")%></font></td>
    </tr>
<%	acres.movenext
    j=j+1
	loop%>
	  <tr bgcolor="#FFFFFF" bordercolor="#CCCCCC"> 
      <td width="14%"> <font face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
        <input type="checkbox" name="chkall" value="checkbox" onclick="javascript:doselectall();">
        </font></td>
      <td colspan="5"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif">Select 
        All </font></td>
    </tr>
	<font face="Verdana,Geneva,Arial,Helvetica,sans-serif"><%call showpageno(pageno)%></font>
	<tr bgcolor="#FFFFFF" bordercolor="#CCCCCC"> 
      <td colspan="6" height="18"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
        <%call showpageno(pageno)
		%>
        </font></td>
    </tr>
  
    
    <tr bgcolor="#FFFFFF" bordercolor="#CCCCCC"> 
      <td colspan="2"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
        <input type="button" name="cmdselect" value="Select" onClick="javascript:dosend();">
        <input type="hidden" name="txtgroupname" value="<%=groupname%>">
        <input type="hidden" name="txtcountry" value="<%=country%>">
        <input type="hidden" name="txtname" value="<%=name%>">
        <input type="hidden" name="txtemail"  value="<%=email%>">
        <input type="hidden" name="txttype" value="<%=strtype%>" >
        <input type="hidden" name="txtcomname" value="<%=companyname%>">
        </font></td>
      <td width="52%"> <font face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
        <input type="hidden" name="txtcomsize" value="<%=companysize%>" >
        <input type="hidden" name="strorder"  value="<%=strorder%>" >
        <input type="hidden" name="radmr" value="<%=sex%>">
        </font> <font face="Verdana,Geneva,Arial,Helvetica,sans-serif">
        <input type="hidden" name="strorderpage" value="<%=strorderpage%>">
        </font></td>
      <td width="52%"> <font face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
        </font></td>
      <td width="52%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif"></font></td>
      <td width="34%"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif"></font></td>
    </tr>
<%else%>
     <tr bgcolor="#FFFFFF" bordercolor="#CCCCCC"> 
      <td colspan="6">No Record</td>
    </tr>
<%end if%>
	
  </table>
</form>
</body>
</html>
<% 
function showpageno(pageno)
	if recount>10 then
		lastpage=recount\10
		if recount mod 10 >0 then
			lastpage=lastpage+1
		end if
		if pageno>10 then
		     response.write "<a href='gselect member.asp?pageno=1&strorder="&strorderpage&"'> The First Page</a>&nbsp;&nbsp;"
			response.write "<a href='gselect member.asp?pageno="&(pageno-9-(pageno  mod 10) )&"&strorder="&strorderpage&"'>Previous 10</a>&nbsp;&nbsp;"
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
				response.write "<a href='gselect member.asp?Pageno="&i&"&strorder="&strorderpage&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;"
			end if	
		next
		if (pageno\10)<(lastpage\10) then
		        response.write "<a href='gselect member.asp?Pageno="&(pageno+11-(pageno mod 10)) &"&strorder="&strorderpage&"'>Next 10</a>&nbsp;&nbsp;"
			   response.write "<a href='gselect member.asp?Pageno="& (lastpage) &"&strorder="&strorderpage&"'> The Last Page</a>&nbsp;&nbsp;"
		end if

		
 end if
end function
%>