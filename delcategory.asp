<!--#include file="database.inc"-->
<%
response.expires=0
if session(("SECURITY_LEVEL"))<>"1" then
				response.redirect "default.asp"
end if
pageno=trim(request("pageno"))
if pageno="" then
  pageno=1
 end if
if session("user_id")="" then
	response.redirect "login.asp"
end if 
straction=request("action_button")
if straction="delete" then
	strid=trim(request("chkid"))
	if strid<>"" then
		strsql="delete from category where id in ("& strid & ")"
		conn.execute strsql
	end if
	response.redirect "delcategory.asp"
end if
%>

<html>
<head>
<title>Enhance Tech</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

function dodelcategory()
{
	var i,j,k,strmsg;
	k=0;
	strmsg="";
	if (document.delcategory.chkid!=null)
	{
		for(i=0;i<document.delcategory.chkid.length;i++)
		{
			if(document.delcategory.chkid[i].checked)
			k++;
		}
		if(i>0 && k==0)
		strmsg="You must select one category to delete!";
		if(i==0)
		{
			if (!document.delcategory.chkid.checked)
			strmsg="You must select one category to delete!";
		}
	}
	if (strmsg!="")
	alert(strmsg);
	else
	{
		document.delcategory.action_button.value="delete";
		document.delcategory.submit();
	}
}
//-->
</script>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" cellspacing="0" cellpadding="0" width="100%" height="400">

  <tr>
    <td width="644" valign="top">
      <div align="center"><a href="DatAdmin.asp">Back 
        to Administrator Management</a> 
        <p>Delete category</p>
        <form name="delcategory" action="delcategory.asp" method="post">
          <table width="66%" border="0" cellspacing="1" cellpadding="4" bgcolor="#C0C0C0">
            <tr bgcolor="#003366"> 
              <td colspan="2" bgcolor="#FFFFFF"> 
                <div align="center"> 
                  <input type="button" name="CmdDelete" value="Delete" onclick="dodelcategory();">
                </div>
              </td>
              <td colspan="2" bgcolor="#FFFFFF"> 
                <div align="center"> 
                  <input type="button" name="CmdBack" value="GoBack" onClick="MM_goToURL('parent','addcategory.asp');return document.MM_returnValue">
                  <input type="hidden" name="action_button" value="">
                </div>
              </td>
            </tr>
            <% 
			  strsql="select * from category order by category"
			 set acres=nothing
			  set acres=createobject("adodb.recordset")
			  acres.cursortype=1
			  acres.locktype=3
			  acres.open strsql,conn
			  recount=acres.recordcount
			  if pageno>1 then
			       i=(pageno-1)*10
				   acres.move i
				end if
			  if acres.eof then
			  %>
            <tr bgcolor="#336699"> 
              <td colspan="4" bgcolor="#FFFFFF">No Record!</td>
            </tr>
            <% Else  %>
            <tr bgcolor="#336699" > 
              <td colspan="4" bgcolor="#FFFFFF"> 
                <% call showpageno(pageno)%>
              </td>
            </tr>
            <%do while not acres.eof
					 if j>9 then exit do 
			  %>
            <tr bgcolor="#336699"> 
              <td width="5%" bgcolor="#FFFFFF"> 
                <input type="checkbox" name="chkid" value="<%= trim(acres("id")&"") %>">
              </td>
              <td colspan="3" bgcolor="#FFFFFF"><%= trim(acres("category")&"") %>&nbsp;</td>
            </tr>
            <%	
			  		acres.movenext
					j=j+1
			  		loop
			   End If %>
            <tr > 
              <td colspan="4" bgcolor="#FFFFFF">
                <% call showpageno(pageno)%>
              </td>
            </tr>
          </table>
        </form>
        </b></div>
    </td>
  </tr>
</table>
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
		     response.write "<a href='delcategory.asp?pageno=1'> The First Page</a>&nbsp;&nbsp;"
			response.write "<a href='delategory.asp?pageno="&(pageno-9-(pageno  mod 10) )&"'>Previous 10</a>&nbsp;&nbsp;"
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
				response.write "<a href='delcategory.asp?Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;"
			end if	
		next
		if (pageno\10)<(lastpage\10) then
		        response.write "<a href='delcategory.asp?Pageno="&(pageno+11-(pageno mod 10)) &"'>Next 10</a>&nbsp;&nbsp;"
			   response.write "<a href='delategory.asp?Pageno="& (lastpage) &"'> The Last Page</a>&nbsp;&nbsp;"
		end if
		
 end if
end function
%>