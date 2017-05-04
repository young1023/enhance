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
		response.redirect "SearchResult.asp"
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
			 response.write strsql
	      conn.execute strsql
		  	 next
	set strsql=nothing
  strsql="delete * from AccList where qty=0 "
		conn.execute strsql
		set strsql=nothing
  		response.redirect "sendto.asp"
 end if
 
if straction="delete" then

  strsql="delete * from AccList where A_id= "& session("enquiryno") & " "
		conn.execute strsql
		set strsql=nothing
		response.redirect "enquiry.asp"
end if
if trim(straction)="" then
	response.redirect "enquiry.asp?action='refresh'&pageno="&pageno&" "
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="imagetoolbar" content="no">
<TITLE>Enhance Technologies Limited</TITLE>
<LINK title=style1 
href="images/basic.css" type=text/css rel=stylesheet><!--BEGIN menuItemData.asp -->
<SCRIPT language=JavaScript 
src="images/js_dynMenuFunctions.js">
<!-- Functions used to create the dynamic HTML drop-down menus -->
</SCRIPT>
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

<!--#include file="menu_bar.js"-->

<!--END menuItemData.asp -->
</HEAD>
<BODY bgColor=#ffffff leftMargin=2 topMargin=2 marginwidth="2" marginheight="2" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<font face="Arial" size="1" color="#FF0000">
<SCRIPT language=javascript>
<!--
//function displayAWindow(url, width, height){
//var Win=window.open(url,"displayAWindow",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=no');
//}

	var gstrBrowserType = null
	if( document.layers )   gstrBrowserType="Netscape";
	else if( document.all )	gstrBrowserType="IE";
	else if( document.getElementById ) gstrBrowserType = "NS6";
	gstrMenuColor = "white"
	gstrMenuColorOn = "#E5DED9"
	gstrBorderColor = "#E41F1F"
	var strAlignment = "left"

	if( gstrBrowserType == "Netscape" )
	{
		style = "gNetMenuText"

		for (var intCnt=0; intCnt<garyMenuNames.length; intCnt++) {
			if( intCnt == (garyMenuNames.length - 1) ){
				//All menus are left-aligned except the last menu
				strAlignment = "left"
				gstrBorderColor = "#4F2403"
				style = "gNetMenuLastText"
			}
			//Pass the menu name, the menu left position Netscape, the Netscape width, and the alignment of text in the menu
			createNetscapeMenu(garyMenuNames[intCnt][0], garyMenuNames[intCnt][1], garyMenuNames[intCnt][3], strAlignment, style)
		}
	}
	else if( gstrBrowserType == "IE" )
	{
		if( "" == "AOL")
		{
			style = "gAOLMenuItemStyle"
		}
		else
		{
			style = "gIEMenuItemStyle"
		}

		for (var intCnt=0; intCnt<garyMenuNames.length; intCnt++) {
			if( intCnt == (garyMenuNames.length - 1) ){
				//All menus are left-aligned except the last menu
				strAlignment = "left"
				gstrBorderColor = "#4F2403"
				style = "gIEMenuLastItemStyle"
			}

			//Pass the menu name, the menu left position for IE, the IE width, and the alignment of text in the menu
			createIEMenu(garyMenuNames[intCnt][0], garyMenuNames[intCnt][2], garyMenuNames[intCnt][4], strAlignment, style)
		}
	}
	else if( gstrBrowserType == "NS6" )
	{
		style = "gIEMenuItemStyle"

		for (var intCnt=0; intCnt<garyMenuNames.length; intCnt++) {
			if( intCnt == (garyMenuNames.length - 1) ){
				//All menus are left-aligned except the last menu
				strAlignment = "left"
				gstrBorderColor = "#4F2403"
				style = "gIEMenuLastItemStyle"
			}

			//Pass the menu name, the menu left position for IE, the IE width, and the alignment of text in the menu
			createNSsixMenu(garyMenuNames[intCnt][0], garyMenuNames[intCnt][5], garyMenuNames[intCnt][6], strAlignment, style)
		}
	}


function openWin(arg_url, arg_name, arg_width, arg_height, arg_scroll, arg_resize) {
    var win = window.open(arg_url,arg_name,'scrollbars=' + arg_scroll + ',resizable=' + arg_resize + ',width=' + arg_width +  ',height=' + arg_height);
  }
//-->
</SCRIPT>
<!--#include file="menu_nav.inc"-->


</SCRIPT>
      
      <TABLE cellSpacing=0 cellPadding=0 width=780 border=0>
        <TR>
          <TD align=center>
<!---- Main Content is inserted here ------>         
            
<table border="0" cellspacing="0" cellpadding="0" width="780" height="400" class="black" >
  
  <tr> 
    
    <td width="644" align="center"> 
      
      <div align="center">
        <form name="form1" action="enquiry.asp" method="post">
          <%	strsql="Select * from Acclist where A_id= "& session("enquiryno") &" ORDER BY Acclist.id DESC  "	
	
	set acres=conn.execute(strsql)				
	if not Acres.eof then 
					%>
          <table width="600" cellspacing="1" border="0" bgcolor="#C0C0C0" class="black" cellpadding="4">
            <tr align="center" > 
              <td colspan="7" height="40" bgcolor="#FFFFFF"> 
                詢價清單</td>
            </tr>
            <tr align="center"  > 
              <td width="171" height="31" bgcolor="#FFFFFF">型號</td>
              <td width="207" height="31" bgcolor="#FFFFFF">數量</td>
              <td width="214" height="31" bgcolor="#FFFFFF">價格範圍</td>
             
            </tr>
            <%  
            	Set Acres = Nothing 
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
            <tr   > 
              <td width="171" bgcolor="#FFFFFF">
              	<% = trim(Acres("P_model")&"") %>
                <input type="hidden" name="txtpm" value="<% = trim(Acres("P_model")&"") %>">
                	</td>
              <td width="207" bgcolor="#FFFFFF">
                <input type="text" name="txtcount" value="<% = Acres("qty") %>" maxlength="5" size="5">
                </td>
              <td width="214" bgcolor="#FFFFFF"> 
                <% if trim(acres("P_price")&"") = "" or trim(acres("P_price")&"") = 0 then %>
              聯絡我們  
                <% else %>
                <% strsql="Select * from exportprice where id="&trim(acres("P_price"))&"  "	  
	               set acr=conn.execute(strsql) 
				   if not acr.eof then
                     response.write mid(trim(acres("p_type")),4)&trim(Acr("exportprice"))
                   else
				%>
                 聯絡我們 
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
            
            <tr align="center" > 
              <td colspan="7" height="14" bgcolor="#FFFFFF"> 
                <input type="button" value="繼續" onClick="dosubmit();" name="button">
                <input type="button" name="Button2" value="完成" onClick="dofin();"> 
                <input type="button" value="更新"  name="button2" onClick="dofresh();"> 
                </font></td>
            </tr>
          </table>
          <%else%>
          <p align="center">你暫時末有選擇貨品。</p> 
          <p align="center">請<a href="SearchResult.asp">繼續選擇</a> !</p> 
          <%end if 
'set strsql=nothing
%>
          <input type="hidden" name="action" value="">
        </form>
      </div>
    </td>
  </tr>
  <tr>
    <td  width="644" valign="top"><font color="#FF0000">&nbsp;</font></td>
  </tr>
</table>
 



<!---- Main Content is ended here ------>  

</TD>
<!--#include file="menu_bottom.inc"-->
</BODY>
<%  
function showpageno(pageno) 
	if recount>10 then 
		lastpage=recount\10 
		if recount mod 10 >0 then 
			lastpage=lastpage+1 
		end if 
		if pageno>10 then 
		     response.write "<a href='enquiry.asp?pageno=1'> The First Page </a>&nbsp;&nbsp;" 
			response.write "<a href='enquiry.asp?pageno="&(pageno-9-(pageno  mod 10) )&"'>Previous 10</a>&nbsp;&nbsp;" 
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
				response.write "<a href='enquiry.asp?Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;" 
			end if	 
		next 
		if (pageno\10)<(lastpage\10) then 
		        response.write "<a href='enquiry.asp?Pageno="&(pageno+11-(pageno mod 10)) &"'>Next 10</a>&nbsp;&nbsp;" 
			   response.write "<a href='enquiry.asp?Pageno="& (lastpage) &"'> The Last Page</a>&nbsp;&nbsp;" 
		end if 
		 
 end if 
end function 
%>