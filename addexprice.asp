<!--#include file="database.inc"-->
<%
response.expires=0 
pageno=trim(request("pageno"))
if pageno="" then
  pageno=1
end if
if session(("SECURITY_LEVEL"))<>"1" then
				response.redirect "default.asp"
end if
if session("user_id")="" then
	response.redirect "login.asp"
end if 

straction=trim(request("action_button"))
strexprice=trim(request("txtexprice"))
if straction="add" then
	if strexprice<>"" then
		strexprice=replace(strexprice,"'","''")
		strsql="Select * from exportprice where exportprice ='"& strexprice & "'"
		set acres=conn.execute(strsql)
		if acres.eof then		
		strsql="insert into exportprice (exportprice) values ('"& strexprice & "')"
		conn.execute strsql
		end if
		set acres=nothing
		response.redirect "addexprice.asp"
	end if
end if
%>
<html>
<head>
<title>Enhance Tech</title>
<LINK title=style1 href="images/basic.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
function doaddexprice()
{
	if(document.addexprice.txtexprice.value=="")
		alert("Please Input Exportprice!");
	else
		{
		document.addexprice.action_button.value="add"
		document.addexprice.submit();
		}
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head>

<body  leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" cellspacing="0" cellpadding="0" width="100%" height="455">
  <tr> 
    <td width="644" height="20">
      </td>
  </tr>
  <tr> 
    <td width="644" valign="top" align="center">
      <a href="datadmin.asp">
       回到管理頁</a></td>
  </tr>
  <tr>
    <td width="644" valign="top">
      <div align="center"><p><b><br>
          折扣管理</b></p>
            <b>
        <p>　</p>
			<form name="addexprice" action="addexprice.asp" method="post">
              
          <table border="0" width="60%" cellspacing="1" cellpadding="4" bgcolor="#C0C0C0">
            <tr bgcolor="#003366"> 
              <td bgcolor="#FFFFFF"><font color="#FFFFCC" size="2" face="Arial, Helvetica, sans-serif"><b>
				Pricing</b></font><font color="#FFFFFF">:</font></td>
              <td bgcolor="#FFFFFF"> 
                <input type="text" name="txtexprice" size="15" maxlength="40">
                </td>
                  
              <td bgcolor="#FFFFFF"> 
                <input type="button" name="CmdAdd" value=" Add "  onClick="doaddexprice()">
                </td>
                  
              <td bgcolor="#FFFFFF"> 
                <input type="reset" name="CmdReset" value="Reset">
                </td>
                  
              <td bgcolor="#FFFFFF"> 
                <input type="button" name="CmdEdit" value=" Edit " onClick="MM_goToURL('parent','Editexprice.asp');return document.MM_returnValue">
				</td>
                  
              <td bgcolor="#FFFFFF"> 
                <input type="button" name="CmDel" value="Delete" onClick="MM_goToURL('parent','Delexprice.asp');return document.MM_returnValue">
                    
                  <input type="hidden" name="action_button" value="">
                 </td>
              </tr>
			  <% 
			  StrSql="select * from exportprice order by exportprice" 
			  set acres=nothing
				set acres=createobject("adodb.recordset")
				acres.cursortype=1
				acres.locktype=3
				acres.open strsql,conn
				'response.write strsql
				recount=acres.recordcount
			
				 if pageno>1 then
				   i=(pageno-1)*10
				   acres.move i
				 end if
			  if acres.eof then
			  %>
                
            <tr bgcolor="#336699"> 
              <td colspan="6" bgcolor="#FFFFFF">No Records</td>
              </tr>
			  <% Else  %>
                
            <tr bgcolor="#003366"> 
              <td colspan="6" bgcolor="#FFFFFF"> 
                <%call showpageno(pageno)%>
              </td>
              </tr>
			  
            <tr bgcolor="#336699" > 
              <td colspan="6" height="34" bgcolor="#FFFFFF"> <font color="#FFFFCC">
				Exist Records:</font> </td>
              </tr>
			  <% 
			  		do while not acres.eof 
					  if j>9 then exit do
			  %>
                
            <tr bgcolor="#336699"> 
              <td colspan="6" bgcolor="#FFFFFF"><%= trim(acres("exportprice")&"") %>&nbsp;%</td>
              </tr>
			  <%
			  		acres.movenext
					j=j+1
			  		loop 
			   End If 
			 %>
			   <tr > 
                  <td colspan="6" bgcolor="#FFFFFF"> <%call showpageno(pageno)%></td>
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
		     response.write "<a href='addexprice.asp?pageno=1'> The First Page</a>&nbsp;&nbsp;"
			response.write "<a href='addexprice.asp?pageno="&(pageno-9-(pageno  mod 10) )&"'>Previous 10</a>&nbsp;&nbsp;"
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
				response.write "<a href='addexprice.asp?Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;"
			end if	
		next
		if (pageno\10)<(lastpage\10) then
		        response.write "<a href='addexprice.asp?Pageno="&(pageno+11-(pageno mod 10)) &"'>Next 10</a>&nbsp;&nbsp;"
			   response.write "<a href='addexprice.asp?Pageno="& (lastpage) &"'> The Last Page</a>&nbsp;&nbsp;"
		end if
		
 end if
end function
%>