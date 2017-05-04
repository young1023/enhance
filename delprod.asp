<!--#include file="database.inc"-->
<%
response.expires=0 
pageno=trim(request("pageno"))
if session(("SECURITY_LEVEL"))<>"1" then
				response.redirect "default.asp"
end if
if session("user_id")="" then
	response.redirect "login.asp"
end if
straction=request("action_button")
if straction="delete" then
	strid=trim(request("chkid"))
	if strid<>"" then
		strsql="delete from product where id in ("& strid & ")"
		conn.execute strsql
		
		response.redirect "delprod.asp"
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
function doedit()
{
	var i,j,k,strmsg;
	k=0;
	strmsg="";
	if (document.editproduct.chkid!=null)
	{
		for(i=0;i<document.editproduct.chkid.length;i++)
		{
			if(document.editproduct.chkid[i].checked)
			k++;
		}
		if(i>0 && k==0)
		strmsg="You must select one or more";
		
		if(i==0)
		{
			if (!document.editproduct.chkid.checked)
			strmsg="You must select one product to edit";
		}
	}
	if (strmsg!="")
	 alert(strmsg);
	else
	{
		if (confirm("Do you want deleting selected products?"))
		{
			document.editproduct.action_button.value="delete";
			document.editproduct.submit();
		}
	}

}


function seachmodel()
{
document.editproduct.action="delprod.asp";
document.editproduct.submit();
}

//-->
</script>

</HEAD>
<BODY >


<div class="container">

<!--#include file="menu_nav.inc"-->

 
      <a href="DatAdmin.asp"><center><p>回到管理頁面</p>
		</center> 
      </a>
          <%
		  pname=trim(request("pname"))
		  if pname="" then
				strsql="SELECT product.description, product.Model, exportprice.Exportprice, category.Category, " _
				      &"product.pic2path, product.id " _
					  &"FROM category RIGHT JOIN (exportprice RIGHT JOIN " _
					  &"product ON exportprice.Id = product.exportprice) " _
					  &" ON category.Id = product.category "
          else
				strsql="SELECT product.description, product.Model, exportprice.Exportprice, category.Category, " _
				      &"product.pic2path, product.id " _
					  &"FROM category RIGHT JOIN (exportprice RIGHT JOIN " _
					  &"product ON exportprice.Id = product.exportprice) " _
					  &" ON category.Id = product.category where product.model like '%"&pname&"%' "
		  end if
				set acres=nothing 
				set acres=createobject("adodb.recordset") 
				acres.cursortype=1 
				acres.locktype=3 
				acres.open strsql,conn 
				recount=acres.recordcount 
				if pageno="" then 
				 pageno=1 
				end if 
				if pageno>1 then 
				  i=(pageno-1)*10 
				  acres.move i 
				 end if  
				if not acres.eof then  
			%>
          </b></p>
        <b><form name="editproduct" action="delprod.asp" method="post">  
              <center>
            <input type="text" name="pname" maxlength="30" size="18" value="<%=pname%>">
            <input type="button" value="Find Model" onclick="seachmodel();">
            <input type="button" name="Submit2" value="Delete selected products" onclick="doedit();">
          </center> 

         <TABLE class="table table-bordered">  

            <tr> 
              <td width="7%" height="22"> 
               
              </td>
              <td width="20%" height="22"> 
                Model
              </td>
              <td width="36%" height="22"> 
                Description
              </td>
              <td width="24%" height="22"> 
                Photo
              </td>
              <td width="13%" height="22"> 
                Price 
                  Range
              </td>
            </tr>
            <%   
			  	do while not acres.eof  
				   if j>9 then exit do 
			  %>
            <tr> 
              <td width="7%" height="89"> 
                <div align="center"> <font face="Arial, Helvetica, sans-serif"> 
                  <input type="checkbox" name="chkid" value="<%= trim(acres("id")&"") %>" >
                  </font></div>
              </td>
              <td width="20%" height="89"><font size="2" face="Arial, Helvetica, sans-serif"><%= trim(acres("model")&"") %>&nbsp;</font></td>
              <td width="36%" height="89"> 
                <div align="center"><font size="2" face="Arial, Helvetica, sans-serif"><%= trim(acres("description")&"") %>&nbsp;</font></div>
              </td>
              <td width="24%" height="89"> <font face="Arial, Helvetica, sans-serif"> 
                <% If not isnull(acres("pic2path")) then %>
                <img src="showimg.asp?id=<%=acres("id")%>&flage=2" width="133" height="100"> 
                <% Else  %>
                No Photo 
                <% End If %>
                </font></td>
              <td width="13%" height="89"> 
                <div align="center"><font face="Arial, Helvetica, sans-serif"><%= trim(acres("exportprice")&"") %>&nbsp;</font></div>
              </td>
            </tr>
            <%  
			  		acres.movenext   
					j=j+1 
			    loop  
			 call showpageno(pageno) %>
          </table>  
			 
			</form>  
			<% Else  %>  
				No Records!
			<% End If %>  


<!--#include file="menu_bottom.inc"-->

</div>

</BODY>

</html>

<%  
function showpageno(pageno) 
	if recount>10 then 
		lastpage=recount\10 
		if recount mod 10 >0 then 
			lastpage=lastpage+1 
		end if 
		if pageno>10 then 
		     response.write "<a href='delprod.asp?pageno=1&pname="&pname&"'> The First Page </a>&nbsp;&nbsp;" 
			response.write "<a href='delprod.asp?pageno="&(pageno-9-(pageno  mod 10) )&"&pname="&pname&"'>Previous 10</a>&nbsp;&nbsp;" 
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
				response.write "<a href='delprod.asp?Pageno="&i&"&pname="&pname&"'>"&cstr(i)&" </a>&nbsp;&nbsp;" 
			end if	 
		next 
		if (pageno\10)<(lastpage\10) then 
		        response.write "<a href='delprod.asp?Pageno="&(pageno+11-(pageno mod 10))&"&pname="&pname&"'>Next 10</a>&nbsp;&nbsp;" 
			   response.write "<a href='delprod.asp?Pageno="& (lastpage) &"&pname="&pname&"'> The Last Page</a>&nbsp;&nbsp;" 
		end if 
		 
 end if 
end function 
%>