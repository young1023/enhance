<%
if session(("SECURITY_LEVEL"))<>"1" then
  response.redirect "default.asp"
end if
session("picflage")=""
session("product_id")=""
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



      <TABLE class="table table-bordered">  
       
 		<tr>
			<td colspan="3" align="center">

			產品管理

             </td>
		
		</tr>
				<tr>
					<td><a href="addprod.asp">
					新增產品</a></td>
			
					<td><a href="product_list.asp">
					修改產品</a></td>
				
					<td><a href="delprod.asp">
					刪除產品
					</a></td>
				</tr>
			</table>

	<TABLE class="table table-bordered">  

		      <tr>
					<td><a href="addcategory.asp">
					新增目錄</a></td>
				
					<td><a href="editcategory.asp">
					修改目錄</a></td>
			
					<td><a href="delcategory.asp">
					刪除目錄</a></td>
				</tr>
			</table>

       
			<TABLE class="table table-bordered"> 

           		
			
		<tr>
			<td colspan="2" align="center">
			查詢管理</td>
		
			<td >
			
			特別優惠</td>
		</tr>
		<tr>
			<td>
		<a href="admin_enquiry.asp">查詢管理</a> 
			</td>
		<td >

						<a href="admin_order.asp">查詢價格管理</a></td>
			<td >
			<a href="sa_offer.asp">優惠項目管理</a> </td>
		</tr>
		</table>     
     
     
 <!--#include file="menu_bottom.inc"-->

</div>

</body>
</html>    