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
<meta name="keywords" content="GC-heat, EOC , Hasco, strack, DME, progressive, Temperature Controller, band heaters, coil heaters, date inserts, Thermocouple, sensor, shot counter, hydraulic cylinder, mold and die components and functional elements, cooling system, industrial heating, vega;wema;hotset;opitz;GC-Heater;�[����;�o����;�o����;�ūױ��;���q��;�[����;�[����;�o����;�o��;�ű��c;�ɰ�;�ɰ��ަ������q;">
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

			���~�޲z

             </td>
		
		</tr>
				<tr>
					<td><a href="addprod.asp">
					�s�W���~</a></td>
			
					<td><a href="product_list.asp">
					�קﲣ�~</a></td>
				
					<td><a href="delprod.asp">
					�R�����~
					</a></td>
				</tr>
			</table>

	<TABLE class="table table-bordered">  

		      <tr>
					<td><a href="addcategory.asp">
					�s�W�ؿ�</a></td>
				
					<td><a href="editcategory.asp">
					�ק�ؿ�</a></td>
			
					<td><a href="delcategory.asp">
					�R���ؿ�</a></td>
				</tr>
			</table>

       
			<TABLE class="table table-bordered"> 

           		
			
		<tr>
			<td colspan="2" align="center">
			�d�ߺ޲z</td>
		
			<td >
			
			�S�O�u�f</td>
		</tr>
		<tr>
			<td>
		<a href="admin_enquiry.asp">�d�ߺ޲z</a> 
			</td>
		<td >

						<a href="admin_order.asp">�d�߻���޲z</a></td>
			<td >
			<a href="sa_offer.asp">�u�f���غ޲z</a> </td>
		</tr>
		</table>     
     
     
 <!--#include file="menu_bottom.inc"-->

</div>

</body>
</html>    