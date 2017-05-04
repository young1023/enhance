<!--#include file="database.inc"-->
<%response.expires=0
pageno=trim(request("pageno"))
strwhere=trim(request("strwhere"))
if pageno="" then
	pageno=1
	strwhere=""
end if


straction=request("action")
function MyDate(strdate)
	stryear=cstr(year(formatdatetime(date,1)))
	if month(date)<10 then
		strmonth="0"& cstr(month(date))
	else
		strmonth=cstr(month(date))
	end if
	if day(date)<10 then
		strday="0"& cstr(day(date))
	else
		strday=cstr(day(date))
	end if
	MyDate=stryear&strmonth&strday
end function
if session("carno")="" then
	strsql="Select Max(id) from AccList"  
		set AcRes=Conn.Execute(strsql)
		 if not Acres.EOF then
		  if trim(acres("expr1000"))<>""then
		    session("carno")=trim(acres("expr1000"))+1
		  else
		    session("carno")=1		
	      end if
		end if
        set AcRes=nothing
		set strsql=nothing
    'RESPONSE.WRITE SESSION("CARNO")
end if 
 
if straction=""  then  
     vpid=split(trim(request("txtpoid")),",")
    if trim(request("txtt"))="" then
		vqty=split(" ",",")
	else
		vqty=split(trim(request("txtt")),",")
   end if
   for i=0 to ubound(vpid)	  
   if trim(vqty(i))<>"" then   	   
         strsql="Select * from product where id="& vpid(i) &" "
				set Acres=Conn.Execute(strsql)		        									
				vprice=Acres("sampleprice")
			  if trim((vprice)&"")="" then
			        vprice=0
			   end if			  
				 vsum=vprice * round(vqty(i))
				 vtype=acres("moneytype")
				vname=replace(trim(Acres("model")&""),"'","''")
                FName=trim(request("FName"))
                LName=trim(request("LName"))
                email=trim(request("email"))
                sd=trim(request("sd1"))+"^ "+trim(request("sd2"))+"^ "+trim(request("sd3"))		
				set strsql=nothing				
		  strsql1="select * from AccList where P_id="& trim(vpid(i))&" and A_id="& session("carno") & " "
		     set Ace=conn.execute(strsql1)
			 if  Ace.eof then
                
                strsql="insert into AccList(A_Id,p_id,qty,P_model,P_Price,p_type,total,saddress) values " _
					  &" ("& session("carno")&"," _
					  &"  "&vpid(i)&"," _
					  &"  "&vqty(i)&"," _
					  &"' "& vname&" '," _
					 &" "&vprice&"," _
					 &" ' "&vtype&" '," _
					  &" "&vsum&",'"&sd&"')"
					 ' response.write strsql
				  Conn.Execute strsql
				 			
		    else
			   totl=round(vqty(i))+Ace("qty")
			   vsum=vprice * totl
			 strsql1="update AccList set A_id= "& session("carno") & " ," _
			          &"P_model='"& vname &"'," _
			          &"P_price="& vprice &"," _
			          &"qty="& totl & "," _
			          &"total="& vsum & "," _
                      &"saddress='"& sd & "' " _
			         &"WHere P_id="& trim(vpid(i))&" and A_id= "& session("carno") & " "
			     conn.execute strsql1
			     set strsql1=nothing 
			end if		 
		  		  
	 	end if  
	next
 	
end if
if straction="save"  then
	vaPid=split(trim(request("txtpid")),",")
	if trim(request("txtcount"))="" then
		vaqty=split(" ",",")
	else
		vaqty=split(trim(request("txtcount")),",")
    end if
	vapm=split(trim(request("txtpm")),",")
	vapc=split(trim(request("txtpc")),",")
	idary=split(trim(request("txtid")),",")
	sd=trim(request("sd1"))+"^ "+trim(request("sd2"))+"^ "+trim(request("sd3"))
	for i=0 to ubound(idary)
    if trim(vaqty(i))="" then	
	     vaqty(i)=0
	  end if
	  vatt=round(trim(vaqty(i))) * trim(vapc(i))	 	  
	strsql="update AccList set A_id= "& session("carno") & " ," _
			  &"P_model='"& replace(trim(vapm(i)),"'","''") &"'," _
			  &"P_price="& trim(vapc(i)) &"," _
			  &"qty="& trim( vaqty(i))& "," _
			  &"total="& trim(vatt) & "," _
  			  &"saddress='"& sd & "' " _
			  &"WHere P_id="& trim(idary(i))&" and A_id= "& session("carno") & " "
			'response.write strsql
	    conn.execute strsql		
	next
	set strsql=nothing
   strsql="delete * from AccList where qty=0 "
		conn.execute strsql
		set strsql=nothing
		response.redirect "SearchResult.asp"
end if

if straction="fresh"  then 
	vaPid=split(trim(request("txtpid")),",")
	if trim(request("txtcount"))="" then
		vaqty=split(" ",",")
	else
		vaqty=split(trim(request("txtcount")),",")
    end if
	vapm=split(trim(request("txtpm")),",")
	vapc=split(trim(request("txtpc")),",")
	idary=split(trim(request("txtid")),",")
	sd=trim(request("sd1"))+"^ "+trim(request("sd2"))+"^ "+trim(request("sd3"))
	for i=0 to ubound(idary)	
	  if trim(vaqty(i)&"")="" then	 
	     vaqty(i)=0
	  end if
	vatt=round(trim(vaqty(i))) * trim(vapc(i))	  
	strsql="update AccList set A_id= "& session("carno") & " ," _
			  &"P_model='"& replace(trim(vapm(i)),"'","''") &"'," _
			  &"P_price="& trim(vapc(i)) &"," _
			  &"qty="& trim( vaqty(i))& "," _
			  &"total="& trim(vatt) & "," _
  			  &"saddress='"& sd & "' " _
			  &"WHere P_id="& trim(idary(i))&" and A_id= "& session("carno") & " "
	  conn.execute strsql	 		
	next
	set strsql=nothing
    strsql="delete * from AccList where qty=0 "
		conn.execute strsql
	set strsql=nothing
	txtt=""
end if


if straction="finorder" then    
    vaPid=split(trim(request("txtpid")),",")
	vaqty=split(trim(request("txtcount")),",")
	vapm=split(trim(request("txtpm")),",")
	vapc=split(trim(request("txtpc")),",")
	vatt=split(trim(request("txttot")),",")
	idary=split(trim(request("txtid")),",")
    session("Name")=trim(request("Name"))
session("cName")=trim(request("cName"))
    session("LName")=trim(request("LName"))
    Session("Email")=trim(request("email"))
    sd=trim(request("sd1"))+", "+trim(request("sd2"))+", "+trim(request("sd3"))


    for i=0 to ubound(idary)	
	   if trim(vaqty(i))="" then	 
	     vaqty(i)=0
		end if
	     vatt=round(trim(vaqty(i))) * trim(vapc(i))	   
	     strsql="update AccList set A_id= "& session("carno") & " ," _
			  &"P_model='"& replace(trim(vapm(i)),"'","''") &"'," _
			  &"P_price="& trim(vapc(i)) &"," _
			  &"qty="& trim( vaqty(i))& "," _
			  &"total="& trim(vatt) & "," _
  			  &"saddress='"& sd & "' " _
			  &"WHere P_id="& trim(idary(i))&" and A_id= "& session("carno") & " "
			 'response.write strsql
	      conn.execute strsql
		  	 next
	set strsql=nothing

  strsql="delete * from AccList where qty=0 "
		conn.execute strsql
		set strsql=nothing


		
  		response.redirect"checkout.asp"
 end if
 
if straction="delete" then

  strsql="delete * from AccList where A_id= "& session("carno") & " "
		conn.execute strsql
		set strsql=nothing
		response.redirect "request.asp"
end if
if trim(straction)="" then
	response.redirect "request.asp?action='refresh'&pageno="&pageno&" "
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
function dofresh()
{
	var i,strmsg,strval,j;
	strmsg="";
	j=0;
	if(document.form1.txtcount != null)
	{
		for(i=0;i<document.form1.txtcount.length;i++)
			{   j=1;
				strval=document.form1.txtcount[i].value;
				
				if(isNaN(strval))
				{
				     strmsg="Please input Number into QTY field.\n";
					break;
				}
				if(strval<0 )
				  {
				  strmsg="please enter a quantity of one or more!.\n";
				  break;
				  }
				if (strval.indexOf(".")!=-1)
				 { 
				 strmsg="please enter a quantity of one or more and must be integer!.\n";
				  break;
				 }
				  
			}
		if (j==0 && i==0)
	    {    
			if(isNaN(document.form1.txtcount.value))
			strmsg="Please input Number into QTY field.\n";
		    if((document.form1.txtcount.value)<0)
		       strmsg="please enter a quantity of one or more!.\n";
		     if (document.form1.txtcount.value.indexOf(".")!=-1)			 
			strmsg="please enter a quantity of one or more and must be integer!.\n";
	      }
	}
	if (strmsg!="")
		alert(strmsg);
	else
	{
		document.form1.action.value="fresh";
		document.form1.submit();
	}
}



function dofin()
{ 

	var i,j,k,strmsg;
	k=0;
	j=0;
	strmsg="";
	for(i=0;i<document.form1.txtcount.length;i++)
		{ 
		 	j=1;
			if(document.form1.txtcount(i).value=="")
			    k++;
			 else 
			  {
				if(isNaN(document.form1.txtcount(i).value))
			      {  strmsg="Please input Number into QTY field.\n";
				  
					break;
				  }
				 if((document.form1.txtcount(i).value)<=0)
				 {
				   strmsg="please enter a quantity of one or more!.\n";
				   break;
				   }
			     if (document.form1.txtcount(i).value.indexOf(".")!=-1)
			      { 
				 strmsg="please enter a quantity of one or more and must be integer!.\n";
				 break;
				 }
			   
		      }
		}
	
	
		if(i>0 && k==document.form1.txtcount.length)
		  strmsg="Please input Number into QTY field.\n";
		
		 if(i==0 &&j==0)
	  {
		     if (document.form1.txtcount.value=="")			    
			  	strmsg="Please input Number into QTY field.\n";
			     else
				  { 
				    if(isNaN(document.form1.txtcount.value))
			           strmsg="Please input Number into QTY field.\n";					  
					  if((document.form1.txtcount.value)<=0)					 
					   strmsg="please enter a quantity of one or more!.\n";			   
			        if (document.form1.txtcount.value.indexOf(".")!=-1)			 
				    strmsg="please enter a quantity of one or more and must be integer!.\n";
			   }
			  
				 
		   }
		   
		    if (document.form1.email.value=="")
	   strmsg="請輸入電郵!";
	   
		   
	  if (document.form1.cname.value=="")
	   strmsg="請輸入公司名稱!";
	   
		    
	if (strmsg!="")
	   alert(strmsg);
	else
	{
		
			document.form1.action.value="finorder";
		document.form1.submit();
	}
}


function dodelect()
{  
		if (confirm("Do you want deleting selected products?"))
		{
			document.form1.action.value="delete";
			document.form1.submit();
		}
		 
}

//-->
</script>

</HEAD>
<BODY>

<div class="container">


<!--#include file="menu_nav.inc"-->


    <form name="form1" action="request.asp" method="post">  

<%	

    strsql="Select * from Acclist where A_id= "& session("carno") &" ORDER BY Acclist.id DESC  "	
	
	set acres=conn.execute(strsql)
				
	if not Acres.eof then
 

%>
		    	   
 <table class="table">


            <tr align="center"   > 

              <td colspan="4" >
 
                詢價清單

              </td>

            </tr>

            <tr align="center"  > 

              <td width="131" height="31" >產品編號</td>

              <td width="153" height="31" >型號</td>

              <td width="154" height="31" >內容</td>

              <td width="152" height="31" >數量</td>

            </tr>
            <%  set acres=nothing 
				  	adopenkeyset=3
					adlockoptimistic=2
					set acres=nothing
					Set acres = createobject("ADODB.Recordset")
					acres.CursorType=adOpenKeyset
					acres.LockType=adLockOptimistic				
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
				j=0		   
		Do while (not acres.eof)
		 if j>9 then exit do 
		%>
		
		 <%
				set w = createobject("adodb.recordset")
           		w.CursorType = adOpenKeyset
                w.LockType=adLockOptimistic
                sql = "SELECT id, Description, category FROM product  where model='"&trim(acres("P_model"))&"'"
               
                w.open sql,conn
                recount=w.recordcount
                if not w.eof then
               %>
            <tr   align="center"> 
              <td width="153" > 
                  <% =trim(w("category")) %>-<% =trim(w("id")) %>
                            </td>
                  <td width="131" > 
              <a href="javascript:;"  onClick="openBrWindow('detail.asp?id=<%=trim(w("id"))%>')" title="detail"><%=trim(Acres("P_model")&"") %></a> 
                <input type="hidden" name="txtpm" value="<%= trim(Acres("P_model")&"") %>">
    
                </font> 
              </td>
              <td width="154" > 
                <% =trim(w("Description")&"") %>
                 <input type="hidden" name="txtpc" value="<%= trim(Acres("P_price")&"")%>">
                  <input type="hidden" name="txttot" value="<%= trim(Acres("total")&"") %>">       
                　</td>
            
              <td width="152" >  
                <input type="text" name="txtcount" value="<%=Acres("qty")%>" maxlength="5" size="5">
              </td>
              <input type="hidden" name="txtid" value="<%= trim(Acres("p_id")&"") %>">
            </tr>
            <%	 end if
				 w.close
			  	 acres.movenext   
				 j=j+1 
			   loop  
			 call showpageno(pageno) 
		     end if
			 set acres=nothing 
			 set sql=nothing
			if sd<>"" then
			sd=split(sd,"^ ")
			sd1=trim(sd(0))
			sd2=trim(sd(1))
			sd3=trim(sd(2))
			end if
		    'for i=0 to ubound(sd)
			%>
 <tr align="center"   > 
              <td height="29"  colspan="4"> 
                <input type="button" value="繼續選擇" onClick="dosubmit();" name="button"> 
                <input type="button" value="更新"  name="button3" onClick="dofresh();"></td>           
            </tr>
 <tr align="center"   > 
              <td height="29"  colspan="4">完成後, 
                
                  請輸入聯絡資料</td>           
            </tr>
 <tr align="center"   > 
              <td height="29" >公司</td>           
              <td height="14" colspan="3" > 
                
                  <input type="text" name="cname" size="67" value="<%=cName%>" maxlength="33">
                 
              </td>
            </tr>
 <tr align="center"   > 
              <td height="29" >姓名</td>  
              <td height="14" colspan="3" > 
               
                  <input type="text" name="Name" size="67" value="<%=Name%>" maxlength="33">
     
              </td>
            </tr>
<tr align="center"   > 
              <td height="29" >電郵</td>
              <td height="14" colspan="3" > 
                
                  <input type="text" name="email" size="67" value="<%=email%>" maxlength="33">
              
              </td>
            </tr>
            <tr align="center"   > 
              <td height="7" colspan="4" > 
                &nbsp;<input type="button" name="Button2" value="完成" onClick="dofin();">&nbsp; 
                </td>
            </tr>
          </table>
          <%else%>
            <TABLE cellSpacing=0 cellPadding=0 width=90% height=300 border=0>
             <tr>
               <td>
          <p align="center"><font size="2" color="#000000" >你暫時末有選擇貨品。</font></p> 
		   <p align="center"><font size="2" color="#000000" >請<a href="SearchResult.asp">繼續選購</a>!</font></p> 
		   </td>
		   </tr>
		   </table> 
<%

end if 

%>

<input type="hidden" name="action" value="">
</form>
 


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
		     response.write "<a href='request.asp?pageno=1'> The First Page </a>&nbsp;&nbsp;" 
			response.write "<a href='request.asp?pageno="&(pageno-9-(pageno  mod 10) )&"'>Previous 10</a>&nbsp;&nbsp;" 
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
				response.write "<a href='request.asp?Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;" 
			end if	 
		next 
		if (pageno\10)<(lastpage\10) then 
		        response.write "<a href='request.asp?Pageno="&(pageno+11-(pageno mod 10)) &"'>Next 10</a>&nbsp;&nbsp;" 
			   response.write "<a href='request.asp?Pageno="& (lastpage) &"'> The Last Page</a>&nbsp;&nbsp;" 
		end if 
		 
 end if 
end function 
%>