<!--#include file="database.inc"-->
<%
if session(("SECURITY_LEVEL"))= "1" then
	response.redirect "DatAdmin.asp"
end if
	stremail=trim(request("txtemail"))
	strpassword=trim(request("txtpassword"))
	
	if stremail<>"" and strpassword<>"" then
		strsql="Select * from member where email='" & stremail & "' and mypassword='"& strpassword & "'"
		response.write strsql
		set acres=conn.execute(strsql)
		if  not acres.eof then
			session("user_id")=trim(acres("id")&"")
			session("SECURITY_LEVEL") = trim(acres("level")&"")
			'session("em")=trim(acres("email")&"")
			session("id")=trim(acres("id")&"")
			session("admin")=trim(acres("level")&"")
			if session(("SECURITY_LEVEL"))="1" then
				response.redirect "DatAdmin.asp"
			else
				strerr="login successully!"
                session(("SECURITY_LEVEL"))="1"
				response.redirect "DatAdmin.asp"
			end if			
		else
			strerr="Login failure. Please login again!"
		end if
	end if
	
%>
<!DOCTYPE html>
<html lang="en">
<HEAD>
<meta charset="Big5"> 
<meta name="keywords" content="GC-heat, EOC , Hasco, strack, DME, progressive, Temperature Controller, band heaters, coil heaters, date inserts, Thermocouple, sensor, shot counter, hydraulic cylinder, mold and die components and functional elements, cooling system, industrial heating, vega;wema;hotset;opitz;GC-Heater;加熱圈;發熱圈;發熱棒;溫度控制器;熱電偶;加熱器;加熱管;發熱筒;油缸;溫控箱;享基;享基科技有限公司;">
<TITLE>Enhance Techologies Limited</TITLE>
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">

<!-- jQuery library -->
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script>

<!-- Latest compiled JavaScript -->
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>

<!-- responsive to mobile devices -->
<meta name="viewport" content="width=device-width, initial-scale=1">

<!--#include file="menu_bar.js"-->

<script language="JavaScript">
<!--

function dologin()
{
	if(document.mainlogin.txtemail.value=="" || document.mainlogin.txtpassword.value=="" )
		alert("Please input login name and password!");
	else
		document.mainlogin.submit();
}

//-->
</script>

</HEAD>
<BODY>

<div class="container">


<!--#include file="menu_nav.inc"-->


<TABLE cellSpacing=0 cellPadding=0 width=98% border=0 height="400">
        <TBODY>
        <TR>
          <TD align=center valign=center>
 
     
    
            <form name="mainlogin" action="login.asp" method="post">
         <TABLE cellPadding=2 border=0 class="normal" width="40%" height="60%">
          <tr> 
            <td height="30">
			
			　</td>
          <td height="30">
              　</td>
        </tr>
          <tr> 
            <td>
			
			　</td>
          <td>
              　</td>
        </tr>
          <tr> 
            <td>
			
			　</td>
          <td>
              　</td>
        </tr>
          <tr> 
            <td>
			
			電郵地址:</td>
          <td>
              <input type="text" name="txtemail" size="18" value="" autocomplete=off>
            </td>
        </tr>
        <tr> 
            <td>
			
			密碼:</td>
          <td >
              <input type="password" name="txtpassword" size="18" value="" autocomplete=off>
            </td>
        </tr>
        <tr> 
            
            <td>&nbsp 
              　</td>
            <td> 
            

              <input type="button" name="cmdSave" value="登入" onclick="dologin()"></td>
        </tr>
        </table>
	  </form>
	  

<!--#include file="menu_bottom.inc"-->

</div>

</BODY>

</html>