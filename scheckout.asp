<!--#include file="database.inc"-->
<%
'on error resume next
response.expires=0
if session("carno")="" then
  response.redirect "searchresult.asp"
end if
stremail=trim(request("txtemail"))
strpassword=trim(request("txtpassword"))
FName=session("FName")
LName=session("LName")
cName=session("cName")
Email=Session("email")

'Response.end
if stremail<>"" and strpassword<>"" then
  strsql="Select * from member where email='" & stremail & "' and mypassword='"& strpassword & "'"
  'response.write strsql
  set acres=conn.execute(strsql)
  if  not acres.eof then
	session("user_id")=trim(acres("id")&"")
	session("SECURITY_LEVEL") = trim(acres("level")&"")
	session("em")=trim(acres("email")&"")
	session("id")=trim(acres("id")&"")
	session("admin")=trim(acres("level")&"")
	saddress=trim(acres("saddress"))			
  else
	strerr="Member login failure, "
  end if
end if	
'If session("user_id")<>"" and session("carno")<>"" then
If session("carno")<>"" then

  strsql="Select * from Acclist  where A_Id="& session("carno") &"   "
  'response.write strsql
  set Acres1=Conn.Execute(strsql)
  if not Acres1.eof then
    strid=session("carno")
	strsql="Select Max(order_no) from AccMain" 
	set AcRes=Conn.Execute(strsql)
	'response.write strsql
	if not Acres.EOF then
	  if trim(acres("expr1000"))<>""then			    
	    orderno=trim(acres("expr1000"))+1
	    'response.write orderno
        'response.end
	  else 
	    orderno=1
	  end if
    end if
    set AcRes=nothing
    set strsql=nothing 

If session("user_id")="" then

   ':: Add Customer Detail into database

     strsql1="insert into Member" _
     & "(Email,CompanyName,FirstName,LastName) values " _
					  &" ('"&Email&"','" _
                       &"  "&cName&"','" _
					  &"  "&FName&"','" _
					  &" "&LName&"')"
					  'response.write strsql1
                      'response.end
				  Conn.Execute strsql1
                 set strsql1=nothing 

     
   ':: Get the Customer ID
      strsql2="Select top 1 * from Member order by id desc" 
	set AcRes2=Conn.Execute(strsql2)

	  CustomerID= AcRes2("ID")

      set AcRes2=nothing
      set strsql2=nothing
  
	'response.write CustomerID
    'response.end	

    strsql="insert into accmain(U_Id,U_No,Order_NO,AccDate) Values(" _
     &" "&CustomerID&", " _
     &" "&session("carno")&" ," _
     &" "& trim(orderno) & " ," _
     &"'"&date& " " & time&"')"
    'response.write strsql
    'response.end

    	End if

If session("user_id")<>"" then


 strsql="insert into accmain(U_Id,U_No,Order_NO,AccDate) Values(" _
     &" "&session("user_id")&", " _
     &" "&session("carno")&" ," _
     &" "& trim(orderno) & " ," _
     &"'"&date& " " & time&"')"
    'response.write strsql
   ' response.end

End if

    Conn.Execute strsql
    set AcRes=nothing
    set strsql=nothing
    strcarno=session("carno")
  else
    strid=session("carno")
  end if
end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=GB2312">
<title>check out</title>
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
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="images/demo/bdg1.jpg" link="#FFFFFF" vlink="ffffcc">
<!--#include file="top.inc"-->
<table border="0" cellspacing="0" cellpadding="0" width="800" height="405">
  <tr> 
    <td valign="top" width="122" height="405" align="center"> 
      <!--#include file="sLeft.inc" -->
    </td>
    <td valign="top" align="center" width="678" height="405"> 
      <!--#include file="sSbar.inc" -->
      <font color="#FFFFFF"><%=strerr%></font> 
     <%'if session("user_id")<>""and session("carno")<>""then 
  if session("carno")<>"" then 
    
      session("carno")=""
		 strsql="Select * from Acclist  where A_Id="& strid &"   "
		 'response.write strsql
		 set Acres1=Conn.Execute(strsql)
		 IsExist=false
  		 if not Acres1.eof then
		  IsExist=true
         end if	
         if IsExist  then
     %>
      <p><font color="#FFFFFF" size="2">Ordering Information</font></p>
      <p><font color="#FFFFFF" size="2">Order Number:</font><font color="#FFFFcc" size="2"><%=orderno%></font></p>
      <p><font color="#FFFFFF" size="2">"Refunds will be given at the discretion of the Company Management"</font></p>
      <form name=w action="https://select.worldpay.com/wcc/purchase" method=POST>
          
        <table width="88%" border="0" cellspacing="1" cellpadding="1" >
          <tr bgcolor="#336699"> 
            <td colspan="4" height="20" align="center"><font color="#FFFF99"><b>Online 
              Payment</b></font></td>
          </tr>
          <tr bgcolor="#336699"> 
            <td width="26%" align="center"> <font color="#FFFF99"><b>Model</b></font> 
            </td>
            <td width="27%" align="center"> <font color="#FFFF99"><b>Number</b></font> 
            </td>
            <td width="23%" align="center"> <font color="#FFFF99"><b>Price</b></font> 
            </td>
            <td width="24%" align="center"> <font color="#FFFF99"><b>Total</b></font> 
            </td>
          </tr>
          <%	
              Do while not AcRes1.eof		
             %>
          <tr bgcolor="#003366"> 
            <td align="center" width="26%"><font color="#FFFFFF"><%= trim(Acres1("p_model") &"")%>&nbsp;</font></td>
            <td align="center" width="27%"> <font color="#FFFFFF"><%= trim(Acres1("qty") &"")%>&nbsp;</font></td>
            <td align="center" width="23%"><font color="#FFFFFF"> 
              <%if trim(acres1("P_price")&"") =0 then
                    p_price=0
                    response.write "<font size=2>Contact Us </font>"
                  else
                    p_price=trim(acres1("p_price"))
				  %>
              <%=trim(acres1("p_type")&"")&trim(Acres1("p_price")&"")%> 
              <%end if%>
              </font></td>
            <td align="center" width="24%"><font color="#FFFFFF"> 
              <%if trim(acres1("total")&"")=0 or trim(acres1("total")&"")="" then %>
              <font size="2">Contact Us</font> 
              <%else
                    yourcount=Cint(acres1("total"))+Cint(yourcount)
                    whatsale=trim(Acres1("p_model"))+","+trim(Acres1("qty")) +","+p_price+","+trim(acres1("total"))+";"+whatsale
				  %>
              <%=trim(acres1("p_type")&"")&trim(Acres1("total")&"")%>&nbsp; 
              <%end if%>
              </font></td>
          </tr>
          <%Acres1.movenext
                whatsale=replace(whatsale,";","<br>")
               loop
			  %>
          <tr bgcolor="#003366"> 
            <td align="center" width="26%"> 
              <input  type=hidden name="instId" value="38019">
              <input type=hidden name="cartId" value="35670">
              <input type=hidden name="testMode" value="0">
              <input type=hidden name="amount" value="<%=yourcount%>">
              <input type=hidden name="desc" value="<%=trim(whatsale)%>">
              <input type=hidden name="currency" value="USD">
              <input type=hidden name="account" value="5609906">
              &nbsp; </td>
            <td align="center" width="27%">&nbsp; </td>
            <td align="center" width="23%">&nbsp;</td>
            <td align="center" width="24%"><font color="#FF6600"><b>USD</b> <%=yourcount%> 
              </font></td>
          </tr>
          <tr bgcolor="#003366"> 
            <td align="center" width="26%" height="16">&nbsp; 
              
            </td>
            <td align="center" colspan="2" height="16">&nbsp;</td>
            <td align="center" width="24%" height="16"><font color="#FFFFFF"> 
              <input type=submit value="Buy It" name="submit">
              </font></td>
          </tr>
          <tr> 
            <td align="center" colspan="4"> 
              <% session("carno")="" %>
              <a href="SearchResult.asp"><font size="3" color="#FFFFCC"><b>Go 
              Back</b></font></a> </td>
          </tr>
          <tr> 
            <td align="center" colspan="4">&nbsp;</td>
          </tr>
        </table>
        </form>
       <%end if
         if IsExist=false then
           session("carno")=""
        %>
           <p align="center"><font color="#ffffff"size="3"><b>This order was delivered,</b></font></p>
          <p align="center"><font color="#ffffff"size="3"><b> please <a href="SearchResult.asp">search and order</a> again !</b></font></p>
        <%End If 
	end if
'end if
 %>
     </div>

   </td>
  </tr>
</table>

</body>
</html>