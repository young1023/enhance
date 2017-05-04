<!--#include file="database.inc"-->
<% 
response.buffer=true
if session(("SECURITY_LEVEL"))<>"1" then
				response.redirect "default.asp"
end if
if session("SECURITY_LEVEL")="" then
	response.redirect "login.asp"
end if
straction=trim(request("action_button"))
if straction="save" then
	strmodel=replace(trim(request("txtmodel")&""),"'","''")
	cstrmodel=replace(trim(request("ctxtmodel")&""),"'","''")
	
	strdes=replace(trim(request("txtdes")),"'","''")
	cstrdes=replace(trim(request("ctxtdes")),"'","''")
	strcid=trim(request("selcategory"))
	strpack=replace(trim(request("txtpack")),"'","''")
	cstrpack=replace(trim(request("ctxtpack")),"'","''")
	strmorder=replace(trim(request("txtminorder")),"'","''")
	strdelivery=replace(trim(request("txtdelivery")),"'","''")
	cstrdelivery=replace(trim(request("ctxtdelivery")),"'","''")
	streprice=trim(request("selexprice"))
	'response.write streprice
	strterm=replace(trim(request("txtterm")),"'","''")
	cstrterm=replace(trim(request("ctxtterm")),"'","''")
	strmoney=trim(request("selmontype"))
	streprice=trim(request("selexprice"))
	strsprice=trim(request("txtsprice"))
	if strsprice="" then
		strsprice="null"
	end if
	if strcid="" then
		strcid="null"
	end if
	if streprice="" then
		streprice="null"
	end if

	strsql="insert into  product( model,cmodel,description,cdescription,category,packing,cpacking,minorder, " _
	      &"delivery,cdelivery,exportprice,term,cterm,sampleprice,moneytype,create_date) " _
		  &"values(" _
	      &"'"& strmodel & "'," _
   	      &"'"& cstrmodel & "'," _
		  &"'"& strdes & "'," _
		  &"'"& cstrdes & "'," _
	      &""& strcid &"," _
		  &"'"& strpack & "'," _
		  &"'"& cstrpack & "'," _
		  &"'"& strmorder & "'," _
		  &"'"& strdelivery & "'," _
		  &"'"& cstrdelivery & "'," _
		  & streprice & ","  _
		  &"'"& strterm & "'," _
		  &"'"& cstrterm & "'," _
		  &""& strsprice & "," _
		  &"'"& strmoney & "'," _
		  &"#"& date & "#)"
	'response.write strsql
	 conn.execute strsql
	 
	 set acresid=conn.execute("select max(id) from product")
	 if not acresid.eof then
	 	session("product_id")=""
	 response.redirect "picupload.asp?product_id="& server.urlencode( trim(acresid.fields(0).value &""))
		'set acresid=nothing
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
function dochkval()
{
	var strmsg;
	strmsg="";
	if(document.addproduct.txtmodel.value=="")
		strmsg="Please input product model.\n";
	if(!document.addproduct.txtsprice.value=="")
		if(isNaN(document.addproduct.txtsprice.value))
			strmsg=strmsg+"Please input sample price as number.\n";
	if(strmsg!="")
		alert(strmsg);
	else
	{
		document.addproduct.action_button.value="save";
		document.addproduct.submit();
	}

}

//-->
</script>
</HEAD>
<BODY>

<div class="container">

<!--#include file="menu_nav.inc"-->
        <form name="addproduct" method="post" action="addprod.asp">  
          
 <TABLE class="table table-bordered">  
   
  <tr>  
    <td colspan="2" align=center>  
        
        <p><a href="DatAdmin.asp">回到管理頁面</a></p>
		<p>新增產品</p>  
      </td>
  </tr>
        
            <tr>   
              <td width="22%" bgcolor="#FFFFFF" >型號</td>  
              <td width="78%" bgcolor="#FFFFFF" >   
                <input type="text" name="txtmodel" size="79" maxlength="100">  
                (required)</td>  
            </tr>  
            <tr>   
              <td width="22%" bgcolor="#FFFFFF" >類型</td>  
              <td width="78%" bgcolor="#FFFFFF" >  
                <select name="selcategory">  
                  <option value="" selected>請選擇類型------------</option>  
                  <% call fillcob("category","") %>  
                </select>  
                </td>  
            </tr>  
            <tr>   
              <td width="22%" valign="top" bgcolor="#FFFFFF" >產品資料</td> 
              <td width="78%" valign="top" bgcolor="#FFFFFF" >
                <textarea name="txtdes" cols="55" rows="5"></textarea> 
                </td> 
            </tr> 
            <tr>  
              <td width="22%" bgcolor="#FFFFFF" >參考資料連結</td> 
              <td width="78%" bgcolor="#FFFFFF" ><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000">  
                <input type="text" name="txtpack" size="79" maxlength="100"> 
                </font></td> 
            </tr> 
            <tr>  
              <td width="22%" bgcolor="#FFFFFF" >
				<font face="Arial, Helvetica, sans-serif" size="2">安裝文件連結</font></td>  
              <td width="78%" bgcolor="#FFFFFF" ><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000">   
                <input type="text" name="txtminorder" size="80" maxlength="100">  
                </font></td>  
            </tr>  
            <tr>   
              <td width="22%" bgcolor="#FFFFFF" >
				<font face="Arial, Helvetica, sans-serif" size="2">技術文件連結</font></td>  
              <td width="78%" bgcolor="#FFFFFF" ><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000">   
                <input type="text" name="txtdelivery" size="80" maxlength="100">  
                </font></td>  
            </tr>  
            <tr>   
              <td width="22%" bgcolor="#FFFFFF" >
				<font face="Arial, Helvetica, sans-serif">價錢範圍</font></td>  
              <td width="78%" bgcolor="#FFFFFF" ><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000">   
                <select name="selexprice">  
                  <option value="" selected>Please select a exportprice---------</option>  
                  <% call fillcob("exprice","") %>  
                </select>  
                </font></td>  
            </tr>  
            <tr>   
              <td width="22%" bgcolor="#FFFFFF" >
				<font face="Arial, Helvetica, sans-serif">條款</font></td>  
              <td width="78%" bgcolor="#FFFFFF" ><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000">   
                <input type="text" name="txtterm" size="81" maxlength="100">  
                </font></td>  
            </tr>  
            <tr>   
              <td width="22%" bgcolor="#FFFFFF" >
				<font face="Arial, Helvetica, sans-serif">所用貨幣</font></td>  
              <td width="78%" bgcolor="#FFFFFF" ><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000">   
                <select name="selmontype" size="1">  
                  <option value="RMB"  selected>RMB</option>  
                  <option value="EURO">EURO</option>
                  <option value="USD">USD </option>  
                  <option value="HKD">HKD</option>  
                </select>  
                </font></td>  
            </tr>  
            <tr>   
              <td width="22%" bgcolor="#FFFFFF" >
				<font face="Arial, Helvetica, sans-serif">零售價錢</font></td>  
              <td width="78%" bgcolor="#FFFFFF" ><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000">   
                <input type="text" name="txtsprice" size="30" maxlength="100">  
                </font></td>  
            </tr>  
            <tr>   
              <td bgcolor="#FFFFFF" >   
                <div align="center"><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000">&nbsp;   
                  </font></div>  
              </td>  
              <td bgcolor="#FFFFFF" ><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000">   
                <input type="button" name="Submit" value="  確定  " onclick="dochkval();">&nbsp;  
                <input type="hidden" name="action_button" value="">  
                </font></td>  
            </tr>  
          </table>  
        </form>  
 

 <!--#include file="menu_bottom.inc"-->

</div>

</body>
</html>