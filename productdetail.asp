<!--#include file="database.inc"-->
<%

strpid=request("id")

session("prid")=strpid



if trim(strpid)<>"" then
	strsqldet="SELECT product.description,product.ID, product.Model, category.Category,  " _
			 &"product.moneytype,product.Pic2Path, " _
			 &"product.packing, product.minorder, product.delivery, exportprice.exportprice, " _
			 &"product.term, product.sampleprice " _
			 &"FROM (product LEFT JOIN category ON product.category = category.Id) " _
			 &"LEFT JOIN exportprice ON product.exportprice = exportprice.Id " _
			 &"Where product.id="& strpid

	set acres=conn.execute(strsqldet)

	if not acres.eof then
	

		strmodel=trim(acres("model")&"")
		strcategory=trim(acres("category")&"")
		strdes=trim(acres("description"))
		strpic=acres("pic2Path")
		strpak=trim(acres("packing")&"")
		strorder=trim(acres("minorder")&"")
		strdelivery=trim(acres("delivery")&"")		
		strexprice=trim(acres("exportprice")&"")				
		strterm=trim(acres("term")&"")						
		strsampleprice=trim(acres("sampleprice")&"")
		strmoneytype=trim(acres("moneytype")&"")
	
	
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
function doorder()
{
	var strmsg;
	var str1;
	strmsg="";	
	if(document.main.txtt.value=="")
	strmsg=strmsg+"Please input a number.\n";
		if(isNaN(document.main.txtt.value))
			strmsg=strmsg+"Please input a number.\n";
		if((document.main.txtt.value)<=0)
		     strmsg="Please input again.\n";
		if (document.main.txtt.value.indexOf(".")!=-1)				 
		strmsg="The number must be an integer.\n";
				
	if(strmsg!="")
		alert(strmsg);
	else
	{
	 
		document.main.submit();
	}

}

//-->
</script>
<SCRIPT language=javascript>
function dosubmit()
{
 var i,j,k,strmsg;
 k=0;
 j=0;
 strmsg="";
 for(i=0;i<document.form1.txtt.length;i++)
  { 
   j=1;
   if(document.form1.txtt(i).value=="")
     k++;
   else 
    {
     if(isNaN(document.form1.txtt(i).value))
       {strmsg="請輸入數量。\n";
         break;
       }
     if((document.form1.txtt(i).value)<=0)
       {strmsg="請輸入數量。\n";
		break;
       }
     if (document.form1.txtt(i).value.indexOf(".")!=-1)
       {strmsg="必須是整數!\n";
        break;
       }
     }
  }

if(i>0 && k==document.form1.txtt.length)
  strmsg="Please input number.\n";

if(i==0 && j==0)
{
 if (document.form1.txtt.value=="")			    
  strmsg="請輸入數量。\n";
 else
  { 
   if(isNaN(document.form1.txtt.value))
     strmsg="請輸入數字。\n";					  
   if((document.form1.txtt.value)<=0)					 
	 strmsg="請輸入正數!\n";			   
   if (document.form1.txtt.value.indexOf(".")!=-1)			 
     strmsg="必須是整數!\n";
  }		  
}

if (strmsg!="")
  alert(strmsg);
else
{
 document.form1.submit();
}

}
//-->
</SCRIPT>

</HEAD>
<BODY>

<div class="container">


<!--#include file="menu_nav.inc"-->



          
                  <input type="hidden" name="txtpoid" value="<%= strpid %>"> 
                
         <table class="table">

            <tr>  

              <td width="434" rowspan="8"  valign="middle" align="center">  
  
                  <%
                  
 PSQL = "Select * From Photo Where PhotoType = 1 and Order_ID ="&acres("id")&" Order by id desc"
                  
                  Set PRs = Conn.Execute(PSQL)
                  
                  If Not PRs.EoF Then
%>  

					<img src="photo/<% = PRs("Path") %>"  border="0">

                  <% Else  %>  
  
                  暫時末有相片  
  
                  <% end if %> 

              </td> 

              <td height="19" align="left" width="65"> 
 
                型號: 

            </td> 

              <td align="left"><%=strmodel%></td> 

            </tr> 

            <tr>  

              <td align="left" width="65"> 
 
                類型:

             </td> 

              <td align="left" width="178">  

				<%=strcategory%></td> 

            </tr> 

            <tr >  

              <td align="left" width="65">  

               內容:

              </td> 

              <td align="left" width="178">  

               <%=strdes%>

              </td> 

            </tr> 

             <tr>  

              <td align="left" width="65">  

               參考資料:

              </td> 

              <td align="left" width="178">

               <% If strpak <> "" Then %>  

               <a target="_blank" href="<% = strpak %>"><font color="red">按此下戴</font></a>

               <% Else %>

               暫末提供

               <% End If%>

              </td> 

            </tr> 

             <tr>  

              <td align="left" width="65">安裝文件</td> 

              <td align="left" width="178">

<% If strorder <> "" Then %>  

              <a href="<% = strorder %>" target=_blank><font color="red">按此下戴</font></a>

 <% Else %>

               暫末提供

               <% End If%>

               　</td> 

            </tr> 

             <tr>  

              <td align="left" width="65">技術文件  

               　</td> 

              <td align="left" width="178">

<% If strdelivery <> "" Then %>  

              <a href="<% = strdelivery %>" target=_blank><font color="red">按此下戴</font></a>

 <% Else %>

               暫末提供

               <% End If%>

               　</td> 

            </tr> 

            <tr>

              <td align="left" width="65"> 
 
         <font color="red">價錢:</font>
           
              </td> 

              <td align="left" width="178">

<% if strexprice="" then%>

                <font color="red">聯絡我們</font>   
 
			<% else %>     

<font color="red"><%=strexprice%></font> 

		    <% end if %> 

               </td> 

            </tr> 

            <tr>

              <td align="left" width="65">  

                <font color="red"  face="Arial">條款:</font> 

              </td> 

              <td align="left" width="178">  
                <font color="red"  face="Arial"><%=strterm%></font> 
              </td> 
            </tr> 
          
            <tr>
              <td width="434" align="center">  
                價格: 
 
                  <% if trim((strsampleprice)&"") = ""  then %>
 
                  聯絡我們  

                  <% else %> 

                  <% = strmoneytype & " " & strsampleprice %> 
  
                  <% end if %>  
                   
              </td> 

              <td width="267" colspan="2"> 
 
                   <% if trim((strsampleprice)&"") = ""  then %>
 
                  要求報價  

                 <%end if%>
               
              </td> 
            </tr> 

  <form name="form1" action="request.asp" method="post"> 
   
        
            <tr>
               
              <td width="434" align="right">數量:</td> 

	  
              <td width="65">  
                  
                 <input type="text" name="txtt" maxlength="4" size="3"> 

              </td>

              <td>

                  <input type="hidden" name="txtpoid" value="<%= trim(Acres("id")&"") %>">

                  <% if trim((strsampleprice)&"") = ""  then %>
 
               <input type="button" value="加入詢價清單" name="B1" onclick="dosubmit()">

                    <% else %> 

                 <input type="button" value="加入購物車" name="B12" onclick="doBuy()">  
  
                 <% end if %>

               	</td> 

            </tr> 
			 
     </table> 
	 
</form> 

                  
<!--#include file="menu_bottom.inc"--> 

</div>

</body>

</html>  