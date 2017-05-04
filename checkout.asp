<!--#include file="database.inc"-->
<%
'on error resume next
response.expires=0

stremail=trim(request("txtemail"))
strpassword=trim(request("txtpassword"))
Name = session("Name")
LName = session("LName")
cName = session("cName")
Email = Session("email")

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

</HEAD>
<BODY>

<div class="container">


<!--#include file="menu_nav.inc"-->
     
<%


  if session("carno")<>"" then 
    
      session("carno")=""
		 strsql="Select * from Acclist a Left Join Product p on a.p_id = p.id where a.A_Id="& strid &"   "
		 'response.write strsql
		 set Acres1=Conn.Execute(strsql)
		 IsExist=false
  		 if not Acres1.eof then
		  IsExist=true
         end if	
         if IsExist  then
     %>

      <p>宣價資料</p>  
      <p>宣問號碼: <%=orderno%></p> 
       
          
        <table class="table">

          <tr  > 

            <td colspan="4" height="20" align="center" bgcolor="#FFFFFF">謝謝你的查詢, 
			我們會把有關的報價發電郵到你所提供的電郵地址。</td>
          </tr>
          <tr  > 
            <td width="26%" align="center" bgcolor="#FFFFFF">型號</td>
            <td width="24%" align="center" bgcolor="#FFFFFF">內容</td>
            <td width="27%" align="center" bgcolor="#FFFFFF">數量</td>
            
          </tr>
          <%	
              Do while not AcRes1.eof		
             %>
          <tr  >
            <td align="center" width="26%" bgcolor="#FFFFFF"><%= trim(Acres1("p_model") &"")%></td>
            <td align="center" width="24%" bgcolor="#FFFFFF"><%= trim(Acres1("description") &"")%></td>
            <td align="center" width="27%" bgcolor="#FFFFFF"><%= trim(Acres1("qty") &"")%></td>
           
          </tr>
          <%Acres1.movenext
                whatsale=replace(whatsale,";","<br>")
               loop
			  %>

          <tr  > 
            <td align="center" width="26%" height="16" bgcolor="#FFFFFF"><font size="1" color="#000000">&nbsp;</font> 
              
            </td>
            <td align="center" height="16" bgcolor="#FFFFFF"></td>
            <td align="center" width="24%" height="16" bgcolor="#FFFFFF">
			<font size="1" color="#000000"> 
              &nbsp;</font></td>
          </tr>
          <tr> 
            <td align="center" colspan="4" bgcolor="#FFFFFF"> 
              <font size="1" color="#000000"> 
              <% session("carno")="" %>
              <a href="default.asp">回到主頁</a></font></td>
          </tr>
          <tr> 
            <td align="center" colspan="4" bgcolor="#FFFFFF"><font size="1" color="#000000">&nbsp;</font></td>
          </tr>
        </table>
        </form>
       <%end if
     
        %>
        <% 
	end if
'end if
 %>



<!--#include file="menu_bottom.inc"-->

</div>

</BODY>

</html>