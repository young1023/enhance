 
<!--#include file="database.inc"-->
<%
		Set SearchConn = Server.CreateObject("ADODB.Connection")
	SearchConn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & server.mappath("database/mydata.mdb")
'	SearchConn.Open "dsn=rental;uid=;pwd="
 %>
<script language="JavaScript">
<!--
function dosearch(p_arg)
{
	var strmsg="";
	
	switch(p_arg){
		case "model":
			if(document.search.txtmodel.value=="")
				strmsg="You must input a word to search.\n";
			break;
		case "category":
			if(document.search.selcategory.value=="-1")
			strmsg="You must select category to search!\n";
			break;
		case "exprice":
			if(document.search.selexprice.value=="-1" )
			strmsg="You must select a export price to search!\n";
			break;
		default:
	}
	if(strmsg!="")
	alert(strmsg);
	else
	{
		document.search.action_button.value=p_arg;
		document.search.submit();
	}
}
function dologin()
{
	if(document.mainlogin.txtemail.value=="" || document.mainlogin.txtpassword.value=="" )
		alert("Please input login email address and password!");
	else
		document.mainlogin.submit();
}

//-->
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="5%">&nbsp;</td>
    <td width="58%" valign="top"> 
      <form name="search" action="searchresult.asp" method="post">
        <table width="96%" border="0" cellspacing="0" cellpadding="0" height="90">
          <tr>
            <td width="41%" height="18"><b><font size="2" color="#EFFBBF" face="Arial, Helvetica, sans-serif">Search 
              By Model:</font></b></td>
            <td width="59%" height="30"><b><font size="2" color="#EFFBBF" face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="txtmodel">
              <a href="javascript:dosearch('model');"><img src="images/demo/bt-next.jpg" width="20" height="20" border="0"></a></font></b></td>
        </tr>
        <tr>
            <td width="41%" height="11"><b><font size="2" color="#EFFBBF" face="Arial, Helvetica, sans-serif">Search 
              By Category:</font></b></td>
            <td width="59%" height="30"><b><font size="2" color="#EFFBBF" face="Arial, Helvetica, sans-serif"> 
              <select name="selcategory">
              <option selected value="-1">Please Select a Category</option>
					<% 
						strsql="select * from Category order by Category"
						set acres=searchconn.execute(strsql)
						do while not acres.eof
					 %>
                    <option value="<%= trim(acres("id")&"") %>"><%= trim(acres("category")&"") %></option>
					<% 
							acres.movenext
						loop
						set acres=nothing
						
					 %>

            </select>
              <a href="javascript:dosearch('category');"><img src="images/demo/bt-next.jpg" width="20" height="20" border="0"></a></font></b></td>
        </tr>
        <tr>
            <td width="41%" height="17"><b><font size="2" color="#EFFBBF" face="Arial, Helvetica, sans-serif">Search 
              By Price:</font></b></td>
            <td width="59%" height="30"><b><font size="2" color="#EFFBBF" face="Arial, Helvetica, sans-serif"> 
              <select name="selexprice">
              <option selected value="-1">Please Select Price </option>
					<% 
						strsql="select * from exportprice order by id"
						set acres=searchconn.execute(strsql)
						do while not acres.eof
					 %>
                    <option value="<%= trim(acres("id")&"") %>"><%= trim(acres("exportprice")&"") %></option>
					<% 
							acres.movenext
						loop
						set acres=nothing
						set searchconn=nothing
					 %>
			  
            </select>
              <a href="javascript:dosearch('exprice');"><img src="images/demo/bt-next.jpg" width="20" height="20" border="0"></a></font></b></td>
			<input type="hidden" name="action_button" value=""> 
        </tr>
      </table>
	</form>
    </td>
    <td width="36%" valign="top"> 
      <% If session("user_id")="" then %>
      <form name="mainlogin" action="checkout.asp" method="post">
        <table border="0" cellspacing="0" cellpadding="0" height="72" width="205">
          <tr> 
            <td><b><font size="2" color="#EFFBBF" face="Arial, Helvetica, sans-serif">E_mail:</font></b></td>
            <td colspan="2" height="30"><b><font size="2" color="#EFFBBF" face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="txtemail" size="18">
              </font></b></td>
        </tr>
        <tr> 
            <td><b><font size="2" color="#EFFBBF" face="Arial, Helvetica, sans-serif">Password:</font></b></td>
            <td colspan="2" height="30"><b><font size="2" color="#EFFBBF" face="Arial, Helvetica, sans-serif"> 
              <input type="password" name="txtpassword" size="18">
              </font></b></td>
        </tr>
        <tr> 
            <td height="30"> 
              <div align="center"><a href="javascript:dologin();"><img src="images/login.gif" width="65" height="25" border="0"></a></div>
          </td>
            <td height="30"> 
              <div align="center"><a href="forgetpwd.asp"><img src="images/forget.gif" width="65" height="25" border="0"></a></div>
          </td>
            <td height="30"> 
              <div align="center"><a href="add_member2.asp"><img src="images/register.gif" width="65" height="25" border="0"></a></div>
          </td>
        </tr>
      </table>
	  </form>
	  <% End If %>&nbsp;
    </td>
    <td width="1%">&nbsp;</td>
  </tr>
</table>
<br>
