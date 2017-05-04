<!--#include file="included/SQLConn.inc.asp"-->
<%
response.expires=0

email = request("email")

If  email <> "" Then

session("SECURITY_LEVEL") = 1
session("user_id") = 1

End If

if session("SECURITY_LEVEL")<>"1" then

response.redirect "default.asp"
				
end if

if session("user_id")="" then
	response.redirect "login.asp"
end if 

pageno=request("pageno")

straction=trim(request("action_button"))


if pageno="" then
	pageno=1
end if

if trim(request("action_button")) = "delete" then

	vid = split(trim(request("chkid")),",")

	for i=0 to ubound(vid)
	
		strsql="Delete Question where QuestionID = "& trim(vid(i))
		
	    conn.execute strsql 
	next
	
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
	document.delorder.action="admin_enquiry.asp";
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
document.delorder.action="admin_enquiry.asp";
document.delorder.submit();
}
// -->
</script>

</HEAD>
<BODY>

<div class="container">

<!--#include file="menu_nav.inc"-->



         
			
			 <% 
			
     	  	sql = "Select * From Question where 1 =1 "
     	  	
     	  	If email <> "" Then

			sql  =  sql & "and Email = '"& email &"'  "

			End If
			 
     	  	sql  =   sql & "Order by QuestionID Desc"			  
			
					adopenkeyset = 1
					adlockoptimistic=2
					set acres=nothing
					Set acres = createobject("ADODB.Recordset")
					acres.CursorType=adOpenKeyset
					acres.LockType=adLockOptimistic
					acres.open sql,conn
					recount=acres.recordcount

			if pageno>1 then  
				i=(pageno-1)*10  
				acres.move i  
			end if  
			if not acres.eof then  
			j=0  
		  %> 
           <br><a href="datadmin.asp">回到管理主頁</a>  /  查詢管理  /  有<%=recount%>條紀錄 
          <form name="delorder" action="admin_enquiry.asp" method="post"> 

            <TABLE class="table table-bordered">  

       <tr >   
                <td bgcolor="#FFFFFF">查詢編號
                </td> 
                <td bgcolor="#FFFFFF">名稱</td>  
                <td bgcolor="#FFFFFF">公司</td>  
                <td bgcolor="#FFFFFF">電郵</td>
                <td bgcolor="#FFFFFF">電話</td>   
                <td bgcolor="#FFFFFF">查詢日期</td>  
                <td bgcolor="#FFFFFF" width="20%">問題</td>
                <td bgcolor="#FFFFFF" >回答</td>
                <td bgcolor="#FFFFFF">回答日期</td>
                <td bgcolor="#FFFFFF">狀況</td>
                <td bgcolor="#FFFFFF">刪除</td>  
              </tr>  
              <%   
	  			do while not acres.eof  
					if j>9 then exit do  
				%>  
              <tr bgcolor="#336699">   
                <td bgcolor="#FFFFFF" ><%= trim(acres("QuestionID")&"") %></td> 
                <td bgcolor="#FFFFFF"><%= trim(acres("Name")&"") %></td> 
                <td bgcolor="#FFFFFF"><%= trim(acres("Company")&"") %></td> 
                <td bgcolor="#FFFFFF"><%= trim(acres("Email")&"") %></td> 
				<td bgcolor="#FFFFFF" ><%= trim(acres("Tel")&"") %></td> 
                <td bgcolor="#FFFFFF" ><% = FormatDateTime(trim(acres("EnquiryDate")&""),2) %></td> 
                <td bgcolor="#FFFFFF"><%= trim(acres("Question")&"") %></td>
                <td bgcolor="#FFFFFF"><% '= trim(acres("")&"") %></td> 
                <td bgcolor="#FFFFFF"><%= trim(acres("ResponseDate")&"") %></td>
                <td bgcolor="#FFFFFF"><%= trim(acres("Status")&"") %></td>   
  			    <td bgcolor="#FFFFFF">   
                <input type="checkbox" name="chkid" value="<%= trim(acres("QuestionID")&"") %>"> 
                                    </td>            
   </tr> 
              <%  
			  		j=j+1 
					acres.movenext 
			  	loop 
			  %> 
              <% If recount>10 then  
				  		lastpage=recount\10 
						if recount mod 10 <>0 then 
							lastpage=lastpage+1 
						end if 
				  %> 
              <tr bgcolor="#336699">  
                <td colspan="11" height="20" bgcolor="#FFFFFF" >  
                  <div align="left">  
                    <% If pageno=1 then %> 
                    <a href="admin_enquiry.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=<%= pageno-1 %>&email=<%=email%>">
					下十條</a>&nbsp;&nbsp; <a href="admin_enquiry.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=<%= lastpage %>&email=<%=email%>">最後十條</a>   
                    <% else    
							if pageno-lastpage=0 then%>  
                    <a href="admin_enquiry.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=<%= lastpage %>&email=<%=email%>">
					最</a><a href="admin_enquiry.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=<%= pageno-1 %>&email=<%=email%>">前十條</a> &nbsp;&nbsp; <a href="admin_enquiry.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=<%= pageno-1 %>&email=<%=email%>">前十條</a>&nbsp;&nbsp;   
                    <%		else%>  
                    <a href="admin_enquiry.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=<%= lastpage %>&email=<%=email%>">
					最</a><a href="admin_enquiry.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=<%= pageno-1 %>&email=<%=email%>">前十條</a> &nbsp;&nbsp; <a href="admin_enquiry.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=<%= pageno-1 %>&email=<%=email%>">前十條</a>&nbsp;&nbsp; 
					<a href="admin_enquiry.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=<%= pageno-1 %>&email=<%=email%>">
					下十條</a>&nbsp;&nbsp; <a href="admin_enquiry.asp?action_button=<%= straction %>&strwhere=<%= strwhere %>&pageno=<%= lastpage %>&email=<%=email%>">最後十條</a>   
                    <%		end if   
						%>  
                    <% End If %>  
                    
</div>  
                </td>  
              </tr>  
              <% end if %>  
              <tr >   
                <td height="32" colspan="5" bgcolor="#FFFFFF">   
                  <div align="left"> 
                  </td>  
                <td height="32" bgcolor="#FFFFFF">   
                  <div align="left">  
                    <input type="button" name="Button" value="Delete" onclick="dodelorder();"> 
                    
</div> 
                </td> 
                <td height="32" colspan="5" bgcolor="#FFFFFF">
				
                </td> 
                <input type="hidden" name="action_button" value=""> 
              </tr> 
            </table>  
		  </form>  
		  <% Else  %>   
                       
            There is no Enquiry Information at the moment.
   
         
            <% End If %>  

      		</td>  
     
  </tr>  
 </table>  
 
 <!--#include file="menu_bottom.inc"-->

</div>

</body>
</html>  