<!--#include file="database.inc"-->
<% 
response.expires=0
if session(("SECURITY_LEVEL"))<>"1" then
				response.redirect "default.asp"
end if
if session("user_id")="" then
	response.redirect "login.asp"
end if 

straction=trim(request("action_button"))
if straction="save" then
	valary=split(trim(request("txtcategory")),",")
	cvalary=split(trim(request("ccategory")),",")
	idary=split(trim(request("txtid")),",")
	for i=0 to ubound(idary)
		strsql="Update category set category='"& trim(replace(valary(i),"'","''"))&"' " _
		      &"where id="& trim(idary(i))
	    conn.execute strsql 
	next
	for i=0 to ubound(idary)
		cstrsql="Update category set ccategory='"& trim(replace(cvalary(i),"'","''"))&"' " _
		      &"where id="& trim(idary(i))
	    conn.execute cstrsql 
	next
end if
%>
<html>
<head>
<title>Enhance Technologies</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

function docategoryedit()
{
	var i,strmsg,strval,j;
	strmsg="";
	j=0;
	if(document.editcategory.txtcategory != null)
	{
		for(i=0;i<document.editcategory.txtcategory.length;i++)
			{
				strval=document.editcategory.txtcategory[i].value;
				if(strval=="")
				{
					strmsg="You must input all data in textbox.";
					j=1;
					break;
				}
			}
		if (j==0 && i==0)
	    {
			if(document.editcategory.txtcategory.value=="")
				strmsg="You must input all data in textbox.";
		}
	}
	if (strmsg!="")
		alert(strmsg);
	else
	{
		document.editcategory.action_button.value="save";
		document.editcategory.submit();
	}
}
//-->
</script>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" cellspacing="0" cellpadding="0" width="100%" height="455">
  <tr>
    <td width="644" valign="top">
      <div align="center"><b>
            
        <p><a href="addCategory.asp">Category Management</a></p>
            <p>Edit category</p>
			<form action="editcategory.asp" method="post" name="editcategory">
              
          <table width="80%" border="0" cellspacing="1" cellpadding="4" bgcolor="#C0C0C0">
            <tr bgcolor="#003366"> 
              <td bgcolor="#FFFFFF"> 
                <div align="center"> 
                    <input type="button" name="cmdSave" value="Save" onclick="docategoryedit()">
                  </div>
                </td>
                  
              <td bgcolor="#FFFFFF"> 
                <div align="center"> 
                    <input type="reset" name="cmdreset" value="Reset">
                  </div>
                </td>
                  
              <td bgcolor="#FFFFFF"> 
                <div align="center"> 
                    <input type="button" name="cmdBack" value="GoBack" onClick="MM_goToURL('parent','addcategory.asp');return document.MM_returnValue">
                    <input type="hidden" name="action_button" value="">
				  </div>
                </td>
              </tr>
			  <% 
			  strsql="Select * from category order by id"
			  set acres=conn.execute(strsql)
			  if acres.eof then
			  %>
                
            <tr bgcolor="336699"> 
              <td colspan="3" bgcolor="#FFFFFF"> No Records </td>
              </tr>
			 <% Else  
			 		Do while not acres.eof
			 %>
                
            <tr bgcolor="336699"> 
              <td colspan="2" bgcolor="#FFFFFF"><% = Trim(Acres("id")) %> 
                <input type="text" name="txtcategory" size="30" maxlength="50" value="<%= trim(acres("category")&"") %>">
                 
               </td>
<td bgcolor="#FFFFFF"> 
                <input type="text" name="ccategory" size="30" maxlength="50" value="<%= trim(acres("ccategory")&"") %>">
                  <input type="hidden" name="txtid" value="<%= trim(acres("id")&"") %>">
               </td>
              </tr>
			 <%
			 		acres.movenext
			 		loop
			  End If %>
            </table>
			</form>
	  
	   </b></div>
    </td>
  </tr>
</table>
</body>
</html>