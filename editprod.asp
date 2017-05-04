<!--#include file="database.inc"-->

<%
if session(("SECURITY_LEVEL"))<>"1" then
				response.redirect "default.asp"
end if
if session("user_id")="" then
	response.redirect "login.asp"
end if
straction=trim(request("action_button"))
strpid=request("product_id")
session("product_id")=strpid
if straction="save" then
	if strpid<>"" then 
		strmodel=replace(trim(request("txtmodel")),"'","''")
        cstrmodel=replace(trim(request("ctxtmodel")),"'","''")
		strdes=replace(trim(request("txtdes")),"'","''")
    	cstrdes=replace(trim(request("ctxtdes")),"'","''")
		strcid=trim(request("selcategory"))
		strpack=replace(trim(request("txtpack")),"'","''")
        cstrpack=replace(trim(request("ctxtpack")),"'","''")
		strmorder=replace(trim(request("txtminorder")),"'","''")
		strdelivery=replace(trim(request("txtdelivery")),"'","''")
       	cstrdelivery=replace(trim(request("ctxtdelivery")),"'","''")
		streprice=trim(request("seleprice"))
		strterm=replace(trim(request("txtterm")),"'","''")
		cstrterm=replace(trim(request("ctxtterm")),"'","''")
        strmoney=trim(request("selmontype"))
		strsprice=trim(request("txtsprice"))
		if strcid="" then
			strcid="null"
		end if
		if streprice="" then
			streprice="null"
		end if
		if strsprice="" then
			strsprice="null"
		end if

		strsql="update product set description='"& strdes & "'," _
			  &"cdescription='"& cstrdes &"'," _
			  &"model='"& strmodel &"'," _
         	  &"cmodel='"& cstrmodel &"'," _
              &"packing='"& strpack &"'," _
              &"cpacking='"& cstrpack &"'," _
			  &"category="& strcid & "," _
			  &"exportprice="& streprice & "," _
			  &"sampleprice="& strsprice & "," _
			  &"delivery='"& strdelivery & "',"  _
			  &"cdelivery='"& cstrdelivery & "',"  _
			  &"term='"& strterm & "'," _
			  &"cterm='"& cstrterm & "'," _
			  &"minorder='"& strmorder & "'," _			  
			  &"moneytype='"& strmoney & "' " _
			  &"WHere id="& strpid
			 ' response.write strsql
	    conn.execute strsql
	end if
	session("product_id")=""
end if
isfind=false
if trim(strpid)<>"" then
	strsql="select * from product where id="& strpid
	set acres=conn.execute(strsql)
	if not acres.eof then
		isfind=true
		strmodel=trim(acres("model")&"")
    	cstrmodel=trim(acres("cmodel")&"")	
        strcid=trim(acres("category")&"")		
		strdes=trim(acres("description")&"")
		cstrdes=trim(acres("cdescription")&"")	
        strpack=trim(acres("packing")&"")
        cstrpack=trim(acres("cpacking")&"")
		strmorder=trim(acres("minorder")&"")
		strdelivery=trim(acres("delivery")&"")
		cstrdelivery=trim(acres("cdelivery")&"")
        streprice=trim(acres("exportprice")&"")
		strterm=trim(acres("term")&"")
    	cstrterm=trim(acres("cterm")&"")
		strmoney=trim(acres("moneytype")&"")
		strsprice=trim(acres("sampleprice")&"")
	end if 
end if
function fillcob(strtype,strid)
	dim acres,strsql,acid,strsel
	select case strtype
		case "category"
			strsql="select id,category as itemval from category order by id"
		case "exprice"
			strsql="select id,exportprice as itemval from exportprice order by exportprice"
	end select
	set acres=conn.execute(strsql)
	do while not acres.eof
		 acid=trim(acres("id")&"")
		 strsel=""
		 if strid<>"" then
			 if acid=strid then
			 	strsel="selected"
			 end if
		 end if
		 response.write "<option value='"&acid&"'"&strsel&">"&trim(acres("itemval")&"")&"</option>"
		 acres.movenext
	loop
	
end function
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

function myopen(what)
{
 window.open('upload_photo.asp?flag='+what,'user','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=20,left=30,width=600,height=400');
}

function opensmall(what)
{
 window.open('upload_small_photo.asp?flag='+what,'user','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=20,left=30,width=600,height=400');
}

function dochkval()
{
	var strmsg;
	strmsg="";
	if(document.editproduct.txtmodel.value=="")
		strmsg="Please input product model.\n";
	if(!document.editproduct.txtsprice.value=="")
		if(isNaN(document.editproduct.txtsprice.value))
			strmsg=strmsg+"Please input sample price as number.\n";
	if(strmsg!="")
		alert(strmsg);
	else
	{
		document.editproduct.action_button.value="save";
		document.editproduct.submit();
	}

}

//-->
</script>

</HEAD>
<BODY>

<div class="container">

<!--#include file="menu_nav.inc"-->

              
<TABLE class="table table-bordered">  

   <tr>   
     <td colspoan="2" align="center">  
      <a href="DatAdmin.asp"><center><p>回到管理頁面</p>
		</center> 
      </a></td>  
  </tr>  
  <tr>  
    <td  valign="top" align="left">  
        
        <% If isfind then %>  
       
  
        <form action="editprod.asp" method="post" name="editproduct">  

         <TABLE class="table table-bordered"> 
            <tr >   
              <td width="26%" bgcolor="#FFFFFF" >型號</td>  
              <td width="74%" bgcolor="#FFFFFF"><font size="2" face="Arial, Helvetica, sans-serif">   
                <input type="text" name="txtmodel" size="56" maxlength="100" value="<%= strmodel %>">  
                (required)</font></td> 
            </tr> 
            <tr >  
              <td width="26%" bgcolor="#FFFFFF" >類型</td> 
              <td width="74%" bgcolor="#FFFFFF"> <font size="2" face="Arial, Helvetica, sans-serif">  
                <select name="selcategory"> 
                  <option value="" <%if strcid="" then%> selected<%end if%>>Please  
                  select a category------------</option> 
                  <% call fillcob("category",strcid) %> 
                </select> 
                </font></td> 
            </tr> 
            <tr >  
              <td width="26%" valign="top" bgcolor="#FFFFFF" >產品資料</td> 
              <td width="74%" valign="top" bgcolor="#FFFFFF" ><font size="2" face="Arial, Helvetica, sans-serif">  
                <textarea name="txtdes" cols="48" rows="5"><%= strdes %></textarea> 
                </font></td> 
            </tr> 
            <tr >  
              <td width="26%" bgcolor="#FFFFFF" >參考資料連結</td> 
              <td width="74%" bgcolor="#FFFFFF"><font size="2" face="Arial, Helvetica, sans-serif">  
                <input type="text" name="txtpack" size="70" maxlength="100" value="<%= strpack %>"> 
                </font></td> 
            </tr> 
            <tr >  
              <td width="26%" bgcolor="#FFFFFF" >
				<font face="Arial, Helvetica, sans-serif" size="2">安裝文件連結</font></td>  
              <td width="74%" bgcolor="#FFFFFF"><font size="2" face="Arial, Helvetica, sans-serif">   
                <input type="text" name="txtminorder" size="70" maxlength="100" value="<%= strmorder %>">  
                </font></td>  
            </tr>  
            <tr >   
              <td width="26%" bgcolor="#FFFFFF" >
				<font face="Arial, Helvetica, sans-serif" size="2">技術文件連結</font></td>  
              <td width="74%" bgcolor="#FFFFFF"><font size="2" face="Arial, Helvetica, sans-serif">   
                <input type="text" name="txtdelivery" size="70" maxlength="100" value="<%= strdelivery %>">  
                </font></td>  
            </tr>  
            <tr >   
              <td width="26%" bgcolor="#FFFFFF" >
				<font face="Arial, Helvetica, sans-serif">折扣</font></td>  
              <td width="74%" bgcolor="#FFFFFF"> <font size="2" face="Arial, Helvetica, sans-serif">   
                <select name="seleprice">  
                  <option value="" <%if streprice="" then %> selected<%end if%>>請選擇折扣</option>  
                  <% call fillcob("exprice",streprice) %>  
                </select>  
                </font></td>  
            </tr>  
            <tr >   
              <td width="26%" bgcolor="#FFFFFF" >
				<font face="Arial, Helvetica, sans-serif">條款</font></td>  
              <td width="74%" bgcolor="#FFFFFF"><font size="2" face="Arial, Helvetica, sans-serif">   
                <input type="text" name="txtterm" size="69" maxlength="100" value="<%= strterm %>">  
                </font></td>  
            </tr> 
            <tr >   
              <td width="26%" bgcolor="#FFFFFF" >
				<font face="Arial, Helvetica, sans-serif">所用貨幣</font></td>  
              <td width="74%" bgcolor="#FFFFFF"> <font size="2" face="Arial, Helvetica, sans-serif">   
                <select name="selmontype" size="1">  
                  <option value="USD" <% If strmoney="USD" then %> selected<% End If %>>USD   
                  </option>  
                  <option value="EURO">EURO</option>
                  <option value="HKD" <% If strmoney="HKD" then %> selected<% End If %>>HKD</option>  
                  <option value="RMB" <% If strmoney="RMB" then %> selected<% End If %>>RMB</option>  
                </select>  
                </font></td>  
            </tr>  
            <tr >   
              <td width="26%" bgcolor="#FFFFFF" >
				<font face="Arial, Helvetica, sans-serif">零售價錢</font><font size="2" face="Arial, Helvetica, sans-serif" > </font></td>  
              <td width="74%" bgcolor="#FFFFFF"><font size="2" face="Arial, Helvetica, sans-serif">   
                <input type="text" name="txtsprice" size="30" maxlength="100" value="<%= strsprice %>">  
                </font> 
				 </td>  
            </tr>  
            <tr >   
              <td width="26%"  height="40" bgcolor="#FFFFFF">   
                修改大相</td>  
              <td width="74%" height="40" bgcolor="#FFFFFF">
              <input type="button" colspan="2"  value="    按此  " onClick="javascript:myopen('<% =strpid %>');"></td>  
            </tr>  
            <tr >   
              <td width="26%"  height="40" bgcolor="#FFFFFF">   
                修改細相</td>  
              <td width="74%" height="40" bgcolor="#FFFFFF">
              <input type="button" colspan="2"  value="    按此  " onClick="javascript:opensmall('<% =strpid %>');">
　</td>  
            </tr>  
            <tr >   
              <td width="26%"  height="40" bgcolor="#FFFFFF">   
                <div align="center">　</div>  
              </td>  
              <td width="74%" height="40" bgcolor="#FFFFFF">  
                <input type="button" name="Submit" value="  確定  " onClick="dochkval();">&nbsp;  
                <input type="hidden" name="action_button" value="">  
                <input type="hidden" name="product_id" value="<%= strpid %>">  
                </td>  
            </tr>  
          </table> 
         </td><td align="center">
         大圖<p>
<%
   Photo_Sql1 = " Select Top 1 * from Photo where Phototype = 1 and Order_id = "&strpid
   Set oPRs = Conn.Execute(Photo_Sql1)
      If Not oPRs.eof then
  do while not oPRs.eof
%>
<a href="Photo/<% = oPRs("Path") %>" target="_blank">
					<img src="Photo/<% = oPRs("Path") %>" width="180" height="135" border="0"></a>
<%                         
                               oPRs.Movenext
							Loop
					else
                      response.write "暫時未有相片"
                    end if
                    
%>
 		</p>
 		細圖<p>&nbsp;<%
   Photo_Sql1 = " Select Top 1 * from Photo where Phototype = 2 and Order_id = "&strpid
   Set oPRs = Conn.Execute(Photo_Sql1)
      If Not oPRs.eof then
  do while not oPRs.eof
%>
<a href="Photo/<% = oPRs("Path") %>" target="_blank">
					<img src="Photo/<% = oPRs("Path") %>" width="180" height="135" border="0"></a>
<%                         
                               oPRs.Movenext
							Loop
					else
                      response.write "暫時未有相片"
                    end if
                    
%>
</td> 
		</form>  
		<% End If %>  
      </div>  
    </td>  
  </tr>  
</table>  
  
</div>
  
<!-- content is ended here ---->  
  
</TD>  
</tr>
</table>
  
  
  
 <!--#include file="menu_bottom.inc"-->

</div>

</body>
</html>  