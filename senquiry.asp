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
if session("enquiryno")="" then
	strsql="Select Max(id) from AccList"  
		set AcRes=Conn.Execute(strsql)
		 if not Acres.EOF then
		  if trim(acres("expr1000"))<>""then
		    session("enquiryno")=trim(acres("expr1000"))+1
		  else
		    session("enquiryno")=1		
	      end if
		end if
        set AcRes=nothing
		set strsql=nothing
    'RESPONSE.WRITE SESSION("enquiryno")
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
				vprice=Acres("exportprice")			
			  if trim((vprice)&"")="" then
			        vprice=0
			   end if				
				vname=replace(trim(Acres("model")&""),"'","''")
				vmty=replace(trim(acres("moneytype")&""),"'","''")			
				set strsql=nothing				
		  strsql1="select * from AccList where P_id="& trim(vpid(i))&" and A_id="& session("enquiryno") & " "
		     set Ace=conn.execute(strsql1)
			 if  Ace.eof then		  
                strsql="insert into AccList (A_Id,p_id,qty,P_model,P_Price,p_type) values " _
					  &" ("& session("enquiryno")&"," _
					  &"  "&vpid(i)&"," _
					  &"  "&vqty(i)&"," _
					  &"' "& vname&" '," _
					  &" "&vprice&"," _
					  &"' "& vmty&" ')"
					'  response.write strsql
				  Conn.Execute strsql
				 			
		    else
			   totl=vqty(i)+Ace("qty")
			 strsql1="update AccList set A_id= "& session("enquiryno") & " ," _
			          &"P_model='"& vname &"'," _
					  &"qty="& totl & "," _
			          &"P_price="& vprice &" " _			        			          
			         &" WHere P_id="& trim(vpid(i))&" and A_id= "& session("enquiryno") & " "
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
	pricetype=split(trim(request("txttype")),",")
	for i=0 to ubound(idary)	
    if trim(vaqty(i)&"")="" then	
	     vaqty(i)=0
	  end if
	  if trim(vapc(i)&"")="" then	
	     vapc(i)="null"
	  end if		 	  
	strsql="update AccList set A_id= "& session("enquiryno") & " ," _
			  &"P_model='"& replace(trim(vapm(i)),"'","''") &"'," _
			  &"P_price="& trim(vapc(i)) &"," _
			  &"P_type='"& trim(pricetype(i)) &"'," _
			  &"qty="& trim( vaqty(i))& " " _			
			  &" WHere P_id="& trim(idary(i))&" and A_id= "& session("enquiryno") & " "
			'response.write strsql
	    conn.execute strsql		
	next
	set strsql=nothing
   strsql="delete * from AccList where qty=0 "
		conn.execute strsql
		set strsql=nothing
		response.redirect "sSearchResult.asp"
end if

if straction="fresh"  then 
	'vaPid=split(trim(request("txtpid")),",")
	if trim(request("txtcount"))="" then
		vaqty=split(" ",",")
	else
		vaqty=split(trim(request("txtcount")),",")
    end if
	vapm=split(trim(request("txtpm")),",")
	vapc=split(trim(request("txtpc")),",")
	idary=split(trim(request("txtid")),",")
	pricetype=split(trim(request("txttype")),",")
	for i=0 to ubound(idary)	
	  if trim(vaqty(i)&"")="" then	 
	     vaqty(i)=0
	  end if
	   if trim(vapc(i)&"")="" then	 
	     vapc(i)=0
	  end if	  
	strsql="update AccList set A_id= "& session("enquiryno") & " ," _
			  &"P_model='"& replace(trim(vapm(i)),"'","''") &"'," _
			  &"P_price="& trim(vapc(i)) &"," _
			  &"P_type='"& trim(pricetype(i)) &"'," _
			  &"qty="& trim( vaqty(i))& "" _			  
			  &" WHere P_id="& trim(idary(i))&" and A_id= "& session("enquiryno") & " "
			  'response.write strsql
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
	pricetype=split(trim(request("txttype")),",")	
	idary=split(trim(request("txtid")),",")	
    for i=0 to ubound(idary)	
	   if trim(vaqty(i))="" then	 
	     vaqty(i)=0
		end if
		if trim(vapc(i))="" then	 
	     vapc(i)=0
		end if	        
	     strsql="update AccList set A_id= "& session("enquiryno") & " ," _
			  &"P_model='"& replace(trim(vapm(i)),"'","''") &"'," _
			  &"P_price="& trim(vapc(i)) &"," _
			   &"P_type='"& trim(pricetype(i)) &"'," _
			  &"qty="& trim( vaqty(i))& " " _			
			  &" WHere P_id="& trim(idary(i))&" and A_id= "& session("enquiryno") & " "
			 'response.write strsql
	      conn.execute strsql
		  	 next
	set strsql=nothing
  strsql="delete * from AccList where qty=0 "
		conn.execute strsql
		set strsql=nothing
  		response.redirect "ssendto.asp"
 end if
 
if straction="delete" then

  strsql="delete * from AccList where A_id= "& session("enquiryno") & " "
		conn.execute strsql
		set strsql=nothing
		response.redirect "senquiry.asp"
end if
if trim(straction)="" then
	response.redirect "senquiry.asp?action='refresh'&pageno="&pageno&" "
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>enguiry</title>
<style type="text/css">
 ul {font-family: "Arial", "Helvetica", "sans-serif"; font-size: 11px; color: #000066; font-weight: bold}
 table {  font-family: "Arial", "Helvetica", "sans-serif"; color: #0000FF; font-size: 11px}
 picposition {  margin-left: 15px; margin-top: 0px; margin-bottom: 0px}</style>
<base target="_self">
<script language="JavaScript">
<!--
function dosubmit()
{
	var i,strmsg,strval,j;
	strmsg="";
	j=0;
	if(document.form1.txtcount != null)
	{
	    
		for(i=0;i<document.form1.txtcount.length ;i++)
			{
			     j=1;
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
		document.form1.action.value="save";
		document.form1.submit();
	}
}



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
		 //if(i>0&& j==document.form1.txtt.length)
		 //strmsg="Please enter a quantity of one or more!.\n";
		
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
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="images/demo/bdg1.jpg" link="ffffff" vlink="ffffcc">
 
<table border="0" cellspacing="0" cellpadding="0" width="100%" height="455">
  <tr align="center"> 
    <td colspan="2" height="125"> 
      <!--#include file="top.inc" -->
    </td>
  </tr>
  <tr> 
    <td valign="top" width="145" align="center" rowspan="3"> 
      <!--#include file="sleft.inc" -->
    </td>
    <td  width="644" valign="top"> 
      <!--#include file="ssbar.inc" -->
      <div align="center">
        <form name="form1" action="senquiry.asp" method="post">
          <%	strsql="Select * from Acclist where A_id= "& session("enquiryno") &" ORDER BY Acclist.id DESC  "	
	
	set acres=conn.execute(strsql)				
	if not Acres.eof then 
					%>
          <table width="600" cellspacing="0" cellpadding="0" border="1">
            <tr align="center" bgcolor="#003366" > 
              <td colspan="7" height="40"> 
                <p><font color="#FFFFFF" size="3" face="Arial, Helvetica, sans-serif"><b><font size="2">报价表</font></b></font></p>
              </td>
            </tr>
            <tr align="center" bgcolor="#003366"> 
              <td width="171" align="center" height="31"><font color="#FFFFFF">型号</font></td>
              <td width="207" align="center" height="31"><font color="#FFFFFF">数量</font></td>
              <td width="214" align="center" height="31"><font color="#FFFFFF">价格范围</font></td>
             
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
            <tr  bgcolor="#003366"> 
              <td align="center" width="171"><font color="#FFFFcc"><%= trim(Acres("P_model")&"") %>&nbsp; 
                <input type="hidden" name="txtpm" value="<%= trim(Acres("P_model")&"") %>">
                </font></td>
              <td align="center" width="207"> <font color="#FFFF99"> 
                <input type="text" name="txtcount" value="<%= Acres("qty") %>" maxlength="5">
                </font></td>
              <td align="center" width="214"><font color="#FFFFcc"> 
                <% if trim(acres("P_price")&"")="" or trim(acres("P_price")&"")=0 then %>
                Contact Us 
                <%else%>
                <% strsql="Select * from exportprice where id="&trim(acres("P_price"))&"  "	
	               set acr=conn.execute(strsql) 
				   if not acr.eof then
                     response.write mid(trim(acres("p_type")),4)&trim(Acr("exportprice"))
                   else
				%>
                Contact Us 
                <%end if
				   end if%>
                <input type="hidden" name="txtpc" value="<%= trim(Acres("P_price")&"")%>">
                </font></td>
              <input type="hidden" name="txtid" value="<%= trim(Acres("p_id")&"") %>">
			  <input type="hidden" name="txttype" value="<%= trim(Acres("p_type")&"") %>">
            </tr>
            <%	 
			  		acres.movenext   
					j=j+1 
			    loop  
			 call showpageno(pageno) 
		     end if
			 set acres=nothing 
			 set sql=nothing
			
			%>
            
            <tr align="center" bgcolor="#003366" > 
              <td colspan="7" height="14"> <font color="#000099"> 
                <input type="button" value="继续购物" onClick="dosubmit();" name="button">
                <input type="button" name="Button2" value="提交" onClick="dofin();">
                <input type="button" value="更新"  name="button2" onClick="dofresh();">
                <input type="button" name="Button" value="清空" onClick="dodelect();">
                </font></td>
            </tr>
          </table>
          <%else%>
          <p align="center"><font color="#ffffff"size="3"><b>你的购物篮没有东西
,</b></font></p>
          <p align="center"><font color="#ffffff"size="3"><b>请继续<a href="sSearchResult.asp">购物</a> !</b></font></p>
          <%end if 
'set strsql=nothing
%>
          <input type="hidden" name="action" value="">
        </form>
      </div>
    </td>
  </tr>
  <tr>
    <td  width="644" valign="top">&nbsp;</td>
  </tr>
</table>
<div align="center"><font size="2" color="#CCCCCC">Copyrightc 2001 Winson Development 
  Company All right reserved.</font> </div>
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
		     response.write "<a href='senquiry.asp?pageno=1'> The First Page </a>&nbsp;&nbsp;" 
			response.write "<a href='senquiry.asp?pageno="&(pageno-9-(pageno  mod 10) )&"'>Previous 10</a>&nbsp;&nbsp;" 
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
				response.write "<a href='senquiry.asp?Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;" 
			end if	 
		next 
		if (pageno\10)<(lastpage\10) then 
		        response.write "<a href='senquiry.asp?Pageno="&(pageno+11-(pageno mod 10)) &"'>Next 10</a>&nbsp;&nbsp;" 
			   response.write "<a href='senquiry.asp?Pageno="& (lastpage) &"'> The Last Page</a>&nbsp;&nbsp;" 
		end if 
		 
 end if 
end function 
%>
