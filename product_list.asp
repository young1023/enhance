<!--#include file="database.inc"-->
<%pageno=trim(request("pageno"))
strwhere=trim(request("strwhere"))
if session(("SECURITY_LEVEL"))<>"1" then
  response.redirect "default.asp"
end if
if pageno="" then
	pageno=1
	strwhere=""
end if%>

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
	if (document.editproduct.product_id!=null)
	{
		for(i=0;i<document.editproduct.product_id.length;i++)
		{
			if(document.editproduct.product_id[i].checked)
			k++;
		}
		if(i>0 && k==0)
		strmsg="You have to select one product";
		if (k>1)
		strmsg="You can only select one product"
		if(i==0)
		{
			if (!document.editproduct.product_id.checked)
			strmsg="You have to select one product to edit";
		}
	}
	if (strmsg!="")
	alert(strmsg);
	else
	{
		document.editproduct.submit();
	}

}

function seachmodel()
{
document.editproduct.action="product_list.asp";
document.editproduct.submit();
}

//-->
</script>
</HEAD>
<BODY>

<div class="container">

<!--#include file="menu_nav.inc"-->


     
    
     <a href="DatAdmin.asp"><center><p>回到管理頁面</p>
		</center> 
      </a>
    
          <%
		  pname=trim(request("pname"))
		  if pname="" then
				strsql  = "SELECT product.description, product.Model, exportprice.Exportprice, category.Category, " _
				      &   " product.pic2path, product.picpath,product.id " _
					  &   " FROM category RIGHT JOIN (exportprice RIGHT JOIN " _
					  &   " product ON exportprice.Id = product.exportprice) " _
					  &   " ON category.Id = product.category "
          else
				strsql="SELECT product.description, product.Model, exportprice.Exportprice, category.Category, " _
				      &"product.pic2path, product.picpath,product.id " _
					  &"FROM category RIGHT JOIN (exportprice RIGHT JOIN " _
					  &"product ON exportprice.Id = product.exportprice) " _
					  &" ON category.Id = product.category where product.model like '%"&pname&"%' "
		  end if
				set acres=nothing 
				'Response.write strsql
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
          </font>
        </p>
        <form name="editproduct" action="editprod.asp" method="post">  
              <center>
            <font color="#FF0000">
            <input type="text" name="pname" maxlength="30" size="18" value="<%=pname%>">
            <input type="button" value="Find Model" onclick="seachmodel();"> 
            <input type="button" name="Submit2" value=" Editing selected product " onclick="doedit();"></font></center>  
          
<TABLE class="table table-bordered">  


            <tr>  
              <td width="5%" bgcolor="#FFFFFF">  
                <div align="center"><b><font color="#FF0000" face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font></b></div> 
              </td> 
              <td width="11%" bgcolor="#FFFFFF">  
                Model
              </td> 
              <td width="27%" bgcolor="#FFFFFF">  
                <div align="center"><b><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000">Description</font></b></div> 
              </td> 
              <td width="20%" bgcolor="#FFFFFF">  
                <div align="center"><b><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000">Small  
                  Photo</font></b></div> 
              </td> 
              <td width="20%" bgcolor="#FFFFFF">  
                <div align="center"><font size="2"><b><font face="Arial, Helvetica, sans-serif" color="#FF0000">Big  
                  Photo</font></b></font></div> 
              </td> 
              <td width="17%" bgcolor="#FFFFFF">  
                <div align="center"><b><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000">Price  
                  Range</font></b></div> 
              </td> 
            </tr> 
            <%    
			  	do while not acres.eof   
				   if j>9 then exit do  
			  %> 
            <tr>  
              <td width="5%" height="94" bgcolor="#FFFFFF">  
                <div align="center"> <font face="Arial, Helvetica, sans-serif" color="#FF0000">  
                  <input type="checkbox" name="product_id" value="<%= trim(acres("id")&"") %>" > 
                  </font></div> 
              </td> 
              <td width="11%" height="94" bgcolor="#FFFFFF"><a href="editprod.asp?product_id=<%= trim(acres("id")&"") %>"><%= trim(acres("model")&"") %></a></td>
              <td width="27%" height="94" bgcolor="#FFFFFF"> 
             <%= trim(acres("description")&"") %>
              </td>
              <td width="20%" height="94" valign="middle" bgcolor="#FFFFFF"> 
                
                
<%
   Photo_Sql1 = " Select Top 1 * from Photo where Phototype = 2 and Order_id = "&acres("id")
   Set oPRs = Conn.Execute(Photo_Sql1)
      If Not oPRs.eof then
  do while not oPRs.eof
%>
<a href="Photo/<% = oPRs("Path") %>" target="_blank">
					<img src="Photo/<% = oPRs("Path") %>" width="100" height="75" border="0"></a>
<%                         
                               oPRs.Movenext
							Loop
					else
                      response.write "no photo"
                    end if
                    
%>
                  
              </td>
              <td width="20%" height="94" bgcolor="#FFFFFF" align="center"> 



<%
   Photo_Sql2 = " Select Top 1 * from Photo where Phototype = 1 and Order_id = "&acres("id")
   Set oPRs1 = Conn.Execute(Photo_Sql2)
      If Not oPRs1.eof then
  do while not oPRs1.eof
%>
<a href="Photo/<% = oPRs1("Path") %>" target="_blank">
<img src="Photo/<% = oPRs1("Path") %>" width="110" height="81" border="0"></a>
<%                         
                               oPRs1.Movenext
							Loop
					else
                      response.write "no photo"
                    end if
                    
%>

              </td>
              <td width="17%" height="94" bgcolor="#FFFFFF"> 
               <%= trim(acres("exportprice")&"") %>
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
		     response.write "<a href='product_list.asp?pageno=1&pname="&pname&"'> The First Page </a>&nbsp;&nbsp;"  
			response.write "<a href='product_list.asp?pageno="&(pageno-9-(pageno  mod 10) )&"&pname="&pname&"'>Previous 10</a>&nbsp;&nbsp;"  
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
				response.write "<a href='product_list.asp?Pageno="&i&"&pname="&pname&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;"  
			end if	  
		next  
		if (pageno\10)<(lastpage\10) then  
		        response.write "<a href='product_list.asp?Pageno="&(pageno+11-(pageno mod 10))&"&pname="&pname&"'>Next 10</a>&nbsp;&nbsp;"  
			   response.write "<a href='product_list.asp?Pageno="& (lastpage) &"&pname="&pname&"'> The Last Page</a>&nbsp;&nbsp;"  
		end if  
		  
 end if  
end function  
%>