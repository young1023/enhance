<!--#include file="included/SQLConn.inc.asp"-->

<%
               
	   
    
	Sql = "Insert into VisitedIP (IPAddress, Page) Values "
	
	Sql = Sql & "('"& Request.ServerVariables("REMOTE_ADDR") &"', 'Special Offer')"
	

        Conn.Execute(Sql)
        
%>

<!--#include file="database.inc"-->
<% 
response.expires=0 

	strsearchsql="SELECT product.description, product.sampleprice, product.offer, product.Model, exportprice.Exportprice, category.Category, " _
	      &"product.pic2path, product.moneytype, product.id " _
		  &"FROM category RIGHT JOIN (exportprice RIGHT JOIN " _
		  &"product ON exportprice.Id = product.exportprice) " _
		  &" ON category.Id = product.category " _
		  &"where product.offer = 1  "



  'response.write strsearchsql

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
       {strmsg="Please input number.\n";
         break;
       }
     if((document.form1.txtt(i).value)<=0)
       {strmsg="Please input again.\n";
		break;
       }
     if (document.form1.txtt(i).value.indexOf(".")!=-1)
       {strmsg="The number must be an integer!.\n";
        break;
       }
     }
  }

if(i>0 && k==document.form1.txtt.length)
  strmsg="Please input number.\n";
  //if(i>0&& j==document.form1.txtt.length)
  //strmsg="Please enter a quantity of one or more!.\n";
if(i==0 && j==0)
{
 if (document.form1.txtt.value=="")			    
  strmsg="Please input number.\n";
 else
  { 
   if(isNaN(document.form1.txtt.value))
     strmsg="Please input number.\n";					  
   if((document.form1.txtt.value)<=0)					 
	 strmsg="Please input again!.\n";			   
   if (document.form1.txtt.value.indexOf(".")!=-1)			 
     strmsg="The number must be an integer !\n";
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

     
 
   			<% If strsearchsql <>"" then %>
            
            <b>                             
            <% 
					i=1
					j=0
					adopenkeyset=3
					adlockoptimistic=2
					set acres=nothing
					Set acres = createobject("ADODB.Recordset")
					acres.CursorType=adOpenKeyset
					acres.LockType=adLockOptimistic
					acres.open strsearchsql,conn
					recount=acres.recordcount
				'response.write recount
					if pageno>1 then
						i=(pageno-1)*10
						acres.move i
					end if
					if not acres.eof then
				%></font>
</b> 
				        
				  <br>本公司現正有 
           <%=recount%> 個產品優惠可供選擇 
            
		   <form name="form1" action="enquiry.asp" method="post">

           
                  <div align="left"> <font size="2"> 
                    <%call showpageno(pageno,straction,strwhere) 
					  oldstrwhere=strwhere
					 %>
                    </font> </div>
              
<table  class="table">
    <tr bgcolor="#FFFFFF">
      <td>相片</td>
<td>型號</td>
<td>內容</td>
<td>價錢</td>
<td>數量</td>
    </tr>
         <tr > 
              <% 	  	do while (not acres.eof)
						if j>8 then exit do
				  %>
             
              <td width = "15%" bgcolor="#FFFFFF">
 <a href="productdetail.asp?id=<%=trim(acres("id")&"") %>"> 
                  
                  <%
   Photo_Sql1 = " Select Top 1 * from Photo where Phototype = 2 and Order_id = "&acres("id")
   Set oPRs = Conn.Execute(Photo_Sql1)
      If Not oPRs.eof then
  do while not oPRs.eof
%>
 <a href="productdetail.asp?id=<%=trim(acres("id")&"")%>">
					<img src="Photo/<% = oPRs("Path") %>" width="100" height="75" border="0"></a>
<%                         
                               oPRs.Movenext
							Loop
					else
                      response.write "no photo"
                    end if
                    
%>
</td>
<td width = "15%" bgcolor="#FFFFFF">

                <a href="productdetail.asp?id=<%=trim(acres("id")&"")%>">
                  
				  <%=acres("model")%></a>


</td>
                 <td  bgcolor="#FFFFFF">
                  <a href="productdetail.asp?id=<%=trim(acres("id")&"")%>">
<%=trim(acres("description")&"") %></a>

</td>
 <td width="13%" bgcolor="#FFFFFF">
            <% = trim(acres("moneytype")&"") %><% = trim(acres("sampleprice")) %>
</td>
<td width="13%" bgcolor="#FFFFFF">
<input type="text" name="txtt" size="3" maxlength="5">
 <input type="hidden" name="txtpoid" value="<%= trim(Acres("id")&"") %>">
       </td>
                 
                
              
              <%  		acres.movenext
						j=j+1 
   %> 
</tr> 
<%				    loop
				  %>

              <tr > 
                <td colspan="5" bgcolor="#FFFFFF"> 
                      <div align="center"> 
                        <font size="2"> 
                        <input type="button" value="Add to Enquiry List" onClick="javascript:dosubmit();" name="button">
                        <%  
						 if trim(session("enquiryno"))<>"" then %>
                        <input type="button" name="Button" value="View List" onClick="javascript:location='enquiry.asp'">  
						  <%  
				          strsql="Select * from Acclist where A_id= "& session("enquiryno") &"  "
				           set acresorder=conn.execute(strsql)				
				          if not Acresorder.eof then %>
				         
                        <input type="button" name="Button2" value="Submit" onClick="javascript:location='sendto.asp'">  
                        <input type="button" name="Button3" value="Empty Basket" onClick="javascript:location='deletecar.asp'">  
                        <%end if  
					       set acresorder=nothing
					      set strsql=nothing
					    %>
                        <input type="hidden" name="hidbuttion" value="">  
                        <%end if%>  
                        </font>
                      </div>
                      <div align="right"> 
                      <font size="2"> 
                      <%call showpageno(pageno,straction,oldstrwhere) %>
                      </font> </div>
                </td>
              </tr>
            </table> 
            </form>				
            <% Else  %>
            暫時末有優惠提供。
            <% End If %>                                     
            <% End If %>                                       
      </td>
    </tr> 
  </table> 
 

 <!--#include file="menu_bottom.inc"-->

</div>

</body>  

</html>
<% 
function showpageno(pageno,straction,strwhere) 
	if recount>9 then 
		lastpage=recount\9 
		if recount mod 9 >0 then 
			lastpage=lastpage+1 
		end if 
		if pageno>9 then 
		     response.write "<a href='searchresult.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&pageno=1'> The First Page</a>&nbsp;&nbsp;" 
			response.write "<a href='searchresult.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&pageno="&(pageno-9-(pageno  mod 9) )&"'>Previous </a>&nbsp;&nbsp;" 
		end if 
		strtemp=pageno 
		if (pageno Mod 9 )=0 then 
		   strtemp=strtemp-9 
		end if 
	 for i=(strtemp-(strtemp mod 9)+1) to (strtemp+9-(strtemp mod 9)) 
	         if lastpage<i then  exit for  
            if i- pageno=0 then 
				response.write cstr(i)&"&nbsp;&nbsp;" 
			else 
				response.write "<a href='searchresult.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;" 
			end if	 
		next 
		if (pageno\9)<(lastpage\9) then 
		       	response.write "<a href='searchresult.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&Pageno="&(pageno+11-(pageno mod 9)) &"'>Next</a>&nbsp;&nbsp;" 
			   response.write "<a href='searchresult.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&Pageno="& (lastpage) &"'> The Last Page</a>&nbsp;&nbsp;" 
		end if 
		 
 end if 
end function 
%>