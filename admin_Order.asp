<!--#include file="database.inc"-->
<%
response.expires=0

email = request("email")

If  email <> "" Then

session("SECURITY_LEVEL") = 1
session("user_id") = 1

End If

if session(("SECURITY_LEVEL"))<>"1" then
				response.redirect "default.asp"
end if
 if session("user_id")="" then
	response.redirect "login.asp"
end if 
pageno=request("pageno")
straction=trim(request("action_button"))
if straction="delete" then
if straction="delete" then
	strid=trim(request("chkid"))
	if strid<>"" then
		strsql="delete from AccList where a_id in ("& strid & ")"
		conn.execute strsql
		strsql="delete from Accmain where U_no in ("& strid & ")"
		conn.execute strsql
	end if
end if

 
	response.redirect "admin_Order.asp"
end if
if pageno="" then
	pageno=1
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

function dodelorder()
{
	var i,j,k,strmsg;
	k=0;
	strmsg="";
	if (document.delorder.chkid!=null)
	{
		for(i=0;i<document.delorder.chkid.length;i++)
		{
			if(document.delorder.chkid[i].checked)
			k++;
		}
		if(i>0 && k==0)
		strmsg="You must select one or more  to delete!";
		if(i==0)
		{
			if (!document.delorder.chkid.checked)
			strmsg="You must select one or more  to delete!";
		}
	}
	if (strmsg!="")
	alert(strmsg);
	else
	{
		document.delorder.action_button.value="delete";
		document.delorder.submit();
	}
}

function seachemail()
{
document.delorder.action="admin_order.asp";
document.delorder.submit();
}
// -->
</script>

</HEAD>

<BODY>

<div class="container">

<!--#include file="menu_nav.inc"-->
       
  
          
			 <%
		 adOpenKeyset=3  
			adLockOptimistic=2  
			email=trim(request("email"))  
			if email="" then  

		  	ordsql="SELECT * " _  
				  & "FROM AccMain INNER JOIN MEMBER ON AccMain.U_Id = MEMBER.ID ORDER BY AccMain.AccDate DESC" 
              
			else  
		  	ordsql="SELECT * " _  
				  &"FROM AccMain INNER JOIN MEMBER ON AccMain.U_Id = MEMBER.ID " _  
				  &" where MEMBER.Email like '%"&email&"%'   ORDER BY AccMain.AccDate DESC "			  
			end if  

           'response.write ordsql

			Set ordres = createobject("ADODB.Recordset")  
			ordres.CursorType = adOpenKeyset  
			ordres.LockType=adLockOptimistic  
			ordres.open ordsql,conn  
			recount=ordres.recordcount  
			'response.write recount  
			if pageno>1 then  
				i=(pageno-1)*10  
				ordres.move i  
			end if  
			if not ordres.eof then  
			j=0  
		  %>  
			  <p>有 <%=recount%> 個紀錄 /   
          查詢價格管理    /   <a href="datadmin.asp">回到管理主頁</a> </p>
 
          <form name="delorder" action="admin_order.asp" method="post"> 
		    <div align=center> 
               
              <input type="text" name="email" maxlength="30" size="23" value="<%=email%>"> 
              <input type="button" value="Email" onclick="seachemail();">  
                
            </div>  

        	<TABLE class="table table-bordered">       
 <tr > 
                <td >   
                 查詢編號 
                </td>   
                
                <td >   
                 內容
                </td> 
               
                   <td >   
                 數量
                </td>
                <td >   
              公司</td> 
                <td >   
              名稱</td> 
                <td >   
              電郵               </td> 
                <td >   
                  查詢日期
                </td>  
                
                  <td >   
                  刪除                </td>
              </tr>  
              <%   
	  			do while not ordres.eof  
					if j>9 then exit do  

                sql1 = "Select * from Acclist  where A_Id= " & trim(ordres("U_no"))

	             set rs1=Conn.Execute(sql1)

				%>  
              <tr > 
                   <td  ><%= trim(ordres("order_no")&"") %></td> 

                <td ><% = rs1("p_model") %></td> 
  
                <td ><% = rs1("qty") %></td> 
  
                 <td  ><%= trim(ordres("CompanyName")&"") %></td> 
 
               
                 <td  ><%= trim(ordres("FirstName")&"") %><%= trim(ordres("LastName")&"") %></td> 
 
                <td  ><%= trim(ordres("Email")&"") %></td> 
                <td  ><%= trim(ordres("AccDate")&"") %></td> 
               <td >   
                  <div align="center">   
                    <input type="checkbox" name="chkid" value="<%= trim(ordres("U_no")&"") %>"> 
                    </div> 
                </td> 
            
   </tr> 
              <%  
			  		j=j+1 
					ordres.movenext 
			  	loop 
			  %> 
              <% If recount>10 then  
				  		lastpage=recount\10 
						if recount mod 10 <>0 then 
							lastpage=lastpage+1 
						end if 
				  %> 
              <tr >  
                <td colspan="7" height="20"  >  
                  <div align="left">  
                    <% If pageno=1 then %> 
                    <a href="admin_Order.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=<%= pageno+1 %>&email=<%=email%>">Next  
                    10</a>&nbsp;&nbsp; <a href="admin_Order.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=<%= lastpage %>&email=<%=email%>">End   
                    10</a>   
                    <% else    
							if pageno-lastpage=0 then%>  
                    <a href="admin_Order.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=1&email=<%=email%>">First   
                    10</a> &nbsp;&nbsp; <a href="admin_Order.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=<%= pageno-1 %>&email=<%=email%>">Prior   
                    10</a>&nbsp;&nbsp;   
                    <%		else%>  
                    <a href="admin_Order.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=1&email=<%=email%> ">First   
                    10</a> &nbsp;&nbsp; <a href="admin_Order.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=<%= pageno-1 %>&email=<%=email%>">Prior   
                    10</a>&nbsp;&nbsp; <a href="admin_Order.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=<%= pageno+1 %>&email=<%=email%>">Next   
                    10</a>&nbsp;&nbsp; <a href="admin_Order.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=<%= lastpage %>&email=<%=email%>">End   
                    10</a>   
                    <%		end if   
						%>  
                    <% End If %>  
                    </div>  
                </td>  
              </tr>  
              <% end if %>  
              <tr >   
                <td height="32" colspan="2" >   
                  <div align="left"></div>  
                  </td>  
                <td height="32" >   
                  　</td> 
                <td height="32" >   
                  　</td> 
                <td height="32" >   
                  <div align="left">  
                    <input type="button" name="Button" value="Delete" onclick="dodelorder();"> 
                    </div> 
                </td> 
                <td height="32" colspan="2" >  
                  <input type="reset" name="Submit2" value=" Reset "> 
                    
                  <div align="left"></div> 
                </td> 
                <input type="hidden" name="action_button" value=""> 
              </tr> 
            </table>  
		  </form>  
		  <% Else  %>   
                       
            There is no Order Information   
         
            <% End If %>  

      	
 
 <!--#include file="menu_bottom.inc"-->

</div>

</body>
</html> 