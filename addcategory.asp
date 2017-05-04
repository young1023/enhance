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
strcategory=trim(request("txtcategory"))
cstrcategory=trim(request("ccategory"))
UpLevel = trim(request("UpLevel"))

if straction="add" then
	if strcategory<>"" then
		strcategory=replace(strcategory,"'","''")
		strsql="Select * from Category where Category ='"& strcategory & "'"
		set acres=conn.execute(strsql)
		if acres.eof then		
		strsql="insert into category (UpLevel, category, ccategory) values ("&UpLevel&", '"& strcategory & "',' "& cstrcategory &"')"
		response.write strsql
		
		conn.execute strsql
		end if
		set acres=nothing
		response.redirect"addcategory.asp"
	end if
end if

function fillcob(strtype,strid)
	dim acres,strsql,acid,strsel
	select case strtype
		case "category"
			strsql="select id,category as itemval from category order by id"
		case "exprice"
			strsql="select id,exportprice as itemval from exportprice order by exportprice"
	end select
	set acres=conn.execute(strsql)
	do while not acres.eof
		 acid=trim(acres("id")&"")
		 strsel=""
		 if strid<>"" then
			 if acid=strid then
			 	strsel="selected"
			 end if
		 end if
		 response.write "<option value='"&acid&"'"&strsel&">"&trim(acres("itemval")&"")&"</option>"
		 acres.movenext
	loop

end function

%>
<html>
<head>
<title>Enhance Technologies</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK title=style1 href="images/basic.css" type=text/css rel=stylesheet><!--BEGIN menuItemData.asp -->

<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
function doaddcategory()
{
	if(document.addcategory.txtcategory.value=="")
		alert("Please Input Category Name!");
	else
		{
		document.addcategory.action_button.value="add"
		document.addcategory.submit();
		}
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head>

<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" cellspacing="1" cellpadding="4" width="100%" height="455">
 
  <tr> 
   
    <td width="644" valign="top" align="center">
          
     <a href="DatAdmin.asp">回到管理頁面</a>
     <br></td>
  </tr>
  <tr>
    <td width="644" valign="top" align="center">
     
	             
        Category Management
            
       <br>
			<form name="addcategory" action="addCategory.asp" method="post">
              
          <table border="0" width="90%" bgcolor="#C0C0C0" cellspacing="1" cellpadding="4">
            <tr bgcolor="#003366"> 
              <td bgcolor="#FFFFFF">
              				上層目錄</td>  
              <td bgcolor="#FFFFFF"> 
                				中文名稱</td>
 <td bgcolor="#FFFFFF"> 
				英文名稱</td>
</tr>
            <tr bgcolor="#003366"> 
              <td bgcolor="#FFFFFF">
              <select name="UpLevel">  
                  <option value="0" selected>請選擇類型------------</option>  
                  <% call fillcob("category","") %>  
                </select>
				　</td>  
              <td bgcolor="#FFFFFF"> 
                <input type="text" name="txtcategory" size="15" maxlength="40">
                </td>
 <td bgcolor="#FFFFFF"> 
                <input type="text" name="ccategory" size="15" maxlength="40">
                </td>
</tr><tr>
                  
            
                  
              <td bgcolor="#FFFFFF"> 
                <p align="center"> 
                <input type="button" name="CmdAdd" value=" Add "  onClick="doaddcategory()" style="float: right"></td>
                  
            
                  
              <td bgcolor="#FFFFFF" width="38%"> 
                <p align="center"> 
                <input type="button" name="CmdEdit" value=" Edit " onClick="MM_goToURL('parent','EditCategory.asp');return document.MM_returnValue">
				</td>
                  
              <td bgcolor="#FFFFFF" width="31%"> 
                <p align="center"> 
                <input type="button" name="CmDel" value="Delete" onClick="MM_goToURL('parent','DelCategory.asp');return document.MM_returnValue">
                    
                  <input type="hidden" name="action_button" value="">
                 </td>
              </tr>
              </table>
			  <% 
			  StrSql="select c1.id as id, c1.category as category, c1.ccategory as ccategory, c2.category as UpLevel from category c1 Left Join category c2 on c1.UpLevel = c2.id order by c1.ID desc" 
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
			  if Not acres.eof then
			  %>
		<br>
		<table cellpadding="4" bgcolor="#C0C0C0" cellspacing="1" width="90%">	  
             
            <tr bgcolor="#003366"> 
              <td colspan="4" bgcolor="#FFFFFF"> 
                <%call showpageno(pageno)%>
              </td>
              </tr>
            <tr bgcolor="#336699" > 
              <td bgcolor="#FFFFFF" width="13%"> ID</td>
              <td  bgcolor="#FFFFFF">上層目錄</td>
              <td  bgcolor="#FFFFFF" width="30%">中文名稱</td>
			<td bgcolor="#FFFFFF" width="36%"> 英文名稱</td>
              </tr>
			  <% 
			  		do while not acres.eof 
					  if j>9 then exit do
			  %>
                
            	<tr bgcolor="#336699"> 
              <td bgcolor="#FFFFFF" width="13%"><%= trim(acres("id")&"") %></td>
              <td bgcolor="#FFFFFF" width="16%"><% = acres("UpLevel") %></td>
              <td bgcolor="#FFFFFF" width="30%"><%= trim(acres("category")&"") %></td>
              <td bgcolor="#FFFFFF"><%= trim(acres("ccategory")&"") %></td>
           
              </tr>
			  <%
			  		acres.movenext
					j=j+1
			  		loop 
			   End If 
			 %>
			   <tr > 
                  <td colspan="4" bgcolor="#FFFFFF"> <%call showpageno(pageno)%></td>
              </tr>
            </table>
            </form>
	  
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
		     response.write "<a href='addcategory.asp?pageno=1'> The First Page</a>&nbsp;&nbsp;"
			response.write "<a href='addcategory.asp?pageno="&(pageno-9-(pageno  mod 10) )&"'>Previous 10</a>&nbsp;&nbsp;"
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
				response.write "<a href='addcategory.asp?Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;"
			end if	
		next
		if (pageno\10)<(lastpage\10) then
		        response.write "<a href='addcategory.asp?Pageno="&(pageno+11-(pageno mod 10)) &"'>Next 10</a>&nbsp;&nbsp;"
			   response.write "<a href='addcategory.asp?Pageno="& (lastpage) &"'> The Last Page</a>&nbsp;&nbsp;"
		end if
		
 end if
end function
%>