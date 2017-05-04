<!--#include file="database.inc"-->
<% 
response.expires=0 
straction=request("action_button")
'response.write straction
if straction<>""then
  session("where")=""
end if

'response.write straction
  if straction=""then
    straction="model"
  end if
stractt=request("action")

pageno=trim(request("pageno"))
strwhere=trim(request("strwhere"))
if pageno="" then
	pageno=1
	strwhere=""
end if
strsearchsql=""
if straction<>"" then
	strsearchsql="SELECT product.description, product.Model, exportprice.Exportprice, category.Category, " _
	      &"product.pic2path, product.id " _
		  &"FROM category RIGHT JOIN (exportprice RIGHT JOIN " _
		  &"product ON exportprice.Id = product.exportprice) " _
		  &" ON category.Id = product.category " _
		  &"where 1 = 1 "
end if
if lcase(straction)="subcategory" then
	strsearchsql="SELECT product.description, product.Model, exportprice.Exportprice, category.Category, " _
	      &"product.pic2path, product.id " _
		  &"FROM category RIGHT JOIN (exportprice RIGHT JOIN " _
		  &"product ON exportprice.Id = product.exportprice)  " _
		  &"ON category.Id = product.category " _
		  &"where 1=1 "
end if
select case lcase(straction)
	case "model"
		strmodel=trim(request("txtmodel"))
		if strmodel<>"" then
			strmodel=replace(strmodel,"'","''")
			strwhere=" and product.model like '%"& strmodel & "%' "
		end if
	case "exprice"
		strexprice=trim(request("selexprice"))
		if strexprice<>"" then
			strwhere=strwhere&" and exportprice.id="& strexprice &" "
		end if
		
	case "category"
	    strcat=trim(request("selcategory"))
		if strcat<>"" then
			strwhere=strwhere & "  and category.id="& strcat &" "
		end if
	case "subcategory"
	    strcat=replace(trim(request("subcategory")),"'","''")
		if strcat<>"" then
			strwhere= "  and category.category like '"& strcat &"%' Order by product.model " 
		end if
end select
if session("where")="" then
  strsearchsql=strsearchsql & strwhere
  session("where")=strwhere
 'response.write session("where")
else 
  strwhere=session("where")
  strsearchsql=strsearchsql & strwhere
end if
 ' response.write strsearchsql


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

	<% If strsearchsql<>"" then %>
            
                                         
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
						i=(pageno-1)*9
						acres.move i
					end if
					if not acres.eof then
				%>

                                             
 
<% 	If straction = "subcategory" then %>

           
             
<div class="row">

<div class="col-sm-2"><h6>找到 <%=recount%> 個紀錄。 </h6></div>

<div class="col-sm-10"><h6><% call showpageno(pageno,straction,strwhere) 

					  oldstrwhere=strwhere

					 %></h6></div>

</div>


<%  End If %>

<div class="row">



<% 	  	

       do while (not acres.eof)

			
             if j > 8 then exit do


' when there is only one record, the display is different

if recount = 1 then  

             desc = acres("description")

%>    
         
        

<% 
		if straction = "category" then
 
%>
  					<a href="searchresult.asp?subcategory=<% = left(trim(acres("model")),10) %>&action_button=subcategory"> 
    
<% 		
		Else    
%>  
 					<a href="productdetail.asp?id=<%=trim(acres("id")&"") %>"> 
 					
<%
		End If
%>
  
<div align="center">   
               
                  <%
                     PSQL = "Select * From Photo Where Order_ID ="&acres("id")&" order by id desc"
                  
                  Set PRs = Conn.Execute(PSQL)
                  
                  If Not PRs.EoF Then
                  
                  
                  %>

               <img src="photo/<% = PRs("Path") %>" border=0 >
                  
                  <%

                    else  

                      response.write "暫時沒有相片"

                    end if

                   %>
                 
                    </a>

<p>

<% 
		if straction = "category" then
		
%>
		
             <a href="searchresult.asp?subcategory=<% = left(trim(acres("model")),10) %>&action_button=subcategory"><% = acres("model") %></a>   
    
<% 		
		Else    
%>                
                <a href="productdetail.asp?id=<%=trim(acres("id")&"")%>"><% = acres("model") %></a>
<%		
		End If
%>				
				
</p>


	         <% = desc %>

</div>	                 	


<% 

		else

			desc = Left(Trim(acres("description")),100)

%>



  		
<% 
		if straction = "category" then
 
%>
  					<a href="searchresult.asp?subcategory=<% = left(trim(acres("model")),10) %>&action_button=subcategory"> 
    
<% 		
		Else    
%>  
 					<a href="productdetail.asp?id=<%=trim(acres("id")&"") %>"> 
 					
<%
		End If
%>
     

<div class="col-sm-4">
            
                  <%
                  
                  PSQL = "Select * From Photo Where Order_ID ="&acres("id")&" order by id desc"
                  
                  Set PRs = Conn.Execute(PSQL)
                  
                  If Not PRs.EoF Then
                  
                  %>

                  <img src="photo/<% = PRs("Path") %>" border=0 >
             
                  <%   

                      else 
 
                      response.write "暫時沒有相片"
                      
                       End If
                  %>
                      </a><p>

                                       
              <% 


		if straction = "category" then
 
%>
  					<a href="searchresult.asp?subcategory=<% = left(trim(acres("model")),10) %>&action_button=subcategory"><% = acres("model") %></a> 
    
<% 	Else  %>  
 					<a href="productdetail.asp?id=<%=trim(acres("id")&"") %>"><% = acres("model") %></a>
 					
               
				  
					
						
<%
		End If
%>   
		
</p>

</div>	
                    

			<% end if %>
 
 	<input type="hidden" name="txtpoid" value="<%= trim(Acres("id")&"") %>">
 					
 						
              
              <%  		
              		acres.movenext

						j=j+1 

                            If j = 3 or j = 6 then 
               %> 

</div>

 <% 
 
 		End If 

 		    Loop
				  %>

    	 
            <% Else  %>

<div class="jumbotron">


<p align="center"> <h3> 暫時沒有紀錄。</h3></p>


</div>     
   
            <% End If %>                                     
            <% End If %>                                       
     
 		

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
		     response.write "<a href='searchresult.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&pageno=1'> First</a>&nbsp;&nbsp;" 
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
		       	response.write "<a href='searchresult.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&Pageno="&(pageno+9-(pageno mod 9)) &"'>Next</a>&nbsp;&nbsp;" 
			   response.write "<a href='searchresult.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&Pageno="& (lastpage) &"'> Last</a>&nbsp;&nbsp;" 
		end if 
		 
 end if 
end function 
%>