<!--#include file="database.inc"-->
<% 
if session(("SECURITY_LEVEL"))<>"1" then
				response.redirect "default.asp"
end if
strid=replace(trim(request("email")),"'","''")
'response.write strid
if strid="" then
	response.redirect "admin_order.asp"
end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>email Detail</title>
<style type="text/css">
 ul {font-family: "Arial", "Helvetica", "sans-serif"; font-size: 11px; color: #000066; font-weight: bold}
 table {  font-family: "Arial", "Helvetica", "sans-serif"; color: #0000FF; font-size: 12px}
 picposition {  margin-left: 15px; margin-top: 0px; margin-bottom: 0px}</style>
<base target="_self">
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
</script>
</head>

<body bgcolor="#FFFFFF" text="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="images/demo/bdg1.jpg" link="#FFFFFF">
<!--#include file="top.inc"-->
<div align="left">

<table border="0" cellspacing="0" cellpadding="0" width="800" height="405">
  <tr>
      <td valign="top" width="122" height="405"><font face="Times New Roman"> 
        <!--#include file="Left.inc" -->
        </font></td>
      <td valign="top" align="center" width="678" height="405"> 
        <div align="center"><!--#include file="sbar.inc" -->
          <div align="center"> 
            <p><a href="admin_Order.asp"><font color="#66FF33" size="2" face="Arial, Helvetica, sans-serif"><b><font size="3">Order 
              Management</font></b></font></a></p>
            <p align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Ordering 
              Information</font></p>
  <p> 
 <%
sqlcmd="select * from member where email= '"& strid &"' "
   set rs=conn.execute(sqlcmd) 
if not rs.eof then     
      pw=rs("mypassword")      
      fn=rs("firstname")
      ln=rs("lastname")
      gd=rs("sex")
      em=rs("email")
      ct=rs("country")
      moble=rs("moble")
      cn=rs("companyname")
      cs=rs("ComanySize")
      bt=rs("businesstype")
	  'saddress=rs("saddress")
	  date1=rs("create_date")
'response.write qt
end if
%>
  </p>
</div>
<div align="center">
            <table border="1" align="center" width="412">
              <tr> 
                <td  width="116" bgcolor="#003366"> 
                  <div align="right"><font color="#FFFFCC" size="2" face="Arial, Helvetica, sans-serif">E_mail 
                    :</font> </div>
                <td  width="280" bgcolor="336699" > 
                  <div align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><%=em%> 
                    </font> </div>
              <tr> 
                <td   width="116" bgcolor="#003366"> 
                  <div align="right"><font color="#FFFFCC" size="2" face="Arial, Helvetica, sans-serif">First 
                    Name: </font> </div>
                <td  width="280" bgcolor="336699"  > 
                  <div align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><%=fn%> 
                    </font> </div>
              <tr> 
                <td   width="116" bgcolor="#003366"> 
                  <div align="right"><font color="#FFFFCC" size="2" face="Arial, Helvetica, sans-serif">Last 
                    Name: </font> </div>
                <td  width="280" bgcolor="336699" > 
                  <div align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><%=ln%> 
                    </font> </div>
             
              <tr> 
                <td   width="116" bgcolor="#003366"> 
                  <div align="right"><font color="#FFFFCC" size="2" face="Arial, Helvetica, sans-serif">Country: 
                    </font> </div>
                <td  width="280" bgcolor="336699"  > 
                  <div align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><%=ct%> 
                    </font> </div>
              <tr> 
                <td   width="116" bgcolor="#003366"> 
                  <div align="right"><font color="#FFFFCC" size="2" face="Arial, Helvetica, sans-serif">Company 
                    Size: </font> </div>
                <td  width="280" bgcolor="336699"  > 
                  <div align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><%=cs%> 
                    </font> </div>
              <tr> 
                <td   width="116" bgcolor="#003366"> 
                  <div align="right"><font color="#FFFFCC" size="2" face="Arial, Helvetica, sans-serif">Company 
                    Name: </font> </div>
                <td  width="280" bgcolor="336699"  > 
                  <div align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><%=cn %> 
                    </font> </div>
              <tr> 
                <td  width="116" bgcolor="#003366"> 
                  <div align="right"><font color="#FFFFCC" size="2" face="Arial, Helvetica, sans-serif">Business 
                    Type: </font> </div>
                <td  width="280" bgcolor="336699"  > 
                  <div align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><%=bt%> 
                    </font> </div>
              <tr> 
                <td   width="116" bgcolor="#003366"> 
                  <div align="right"><font color="#FFFFCC" size="2" face="Arial, Helvetica, sans-serif">Register 
                    date:</font></div>
                <td  width="280" bgcolor="336699"  > 
                  <div align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><%=date1%></font></div>
            </table>  
   </div>

          
    </div>
  </tr>
</table>
</div>
</body>
</html>
