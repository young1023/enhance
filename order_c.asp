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
		response.redirect "SearchResult_c.asp"
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
    session("FName")=trim(request("FName"))
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
  		response.redirect"checkout_c.asp"
 end if
 
if straction="delete" then

  strsql="delete * from AccList where A_id= "& session("carno") & " "
		conn.execute strsql
		set strsql=nothing
		response.redirect "order_c.asp"
end if
if trim(straction)="" then
	response.redirect "order_c.asp?action='refresh'&pageno="&pageno&" "
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="imagetoolbar" content="no">
<TITLE>Jaguar Pen Limited</TITLE>
<LINK title=style1 
href="images/basic.css" type=text/css rel=stylesheet><!--BEGIN menuItemData.asp -->
<SCRIPT language=JavaScript 
src="images/js_dynMenuFunctions.js">
<!-- Functions used to create the dynamic HTML drop-down menus -->
</SCRIPT>
<script language="JavaScript">
<!--
function openBrWindow(theURL) {
  window.open(theURL, 'basket','menubar=no,toolbar=no,location=no,directories=no,status=no,scrollbars=1,resizable=0,width=400,top=80,left='+(window.screen.availWidth-500)+',height=360');
}

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
    if (document.form1.sd1.value=="")
	   strmsg="Please input Shipping Address !";

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


<!--#include file="menu_bar_c.js"-->

<!--END menuItemData.asp -->
</HEAD>
<BODY leftMargin=2 topMargin=2 marginwidth="2" marginheight="2" text="#FF0000" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
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
<!--#include file="menu_nav_c.inc"-->

<SCRIPT language=JavaScript>
<!--
var dcs_imgarray = new Array;
var dcs_ptr = 0;
var dCurrent = new Date();
var DCS=new Object();
var WT=new Object();
var DCSext=new Object();
var saveReqArgs = "";

//var dcsADDR = "64.15.240.148";
var dcsADDR = "metric.organic.com";
var dcsID = "";

if (dcsID == ""){
	var TagPath = dcsADDR;
} else {
	var TagPath = dcsADDR+"/"+dcsID;
}

function dcs_var(){
	WT.tz = dCurrent.getTimezoneOffset();
	WT.ul = navigator.appName=="Netscape" ? navigator.language : navigator.userLanguage;
	WT.cd = screen.colorDepth;
	WT.sr = screen.width+"x"+screen.height;
	WT.jo = navigator.javaEnabled() ? "Yes" : "No";
	WT.ti   = document.title;
	DCS.dcsdat = dCurrent.getTime();
	if ((window.document.referrer != "") && (window.document.referrer != "-")){
		if (!(navigator.appName == "Microsoft Internet Explorer" && parseInt(navigator.appVersion) < 4) ){
			DCS.dcsref = window.document.referrer;
		}
	}
    	if ( !window.forcePathname) {
	    DCS.dcsuri = window.location.pathname;
	    //alert ('window.location.pathname = ' + window.location.pathname);
	} else {
	    DCS.dcsuri = forcePathname;
	    //alert ('forcePathname = ' + forcePathname);
	}
	DCS.dcsqry = window.location.search;
	DCS.dcssip = window.location.hostname;
}

function A(N,V){
	return "&"+N+"="+escape(V);
}

function dcs_createImage(dcs_src)
{
	if (document.images){
		dcs_imgarray[dcs_ptr] = new Image;
		dcs_imgarray[dcs_ptr].src = dcs_src;
		dcs_ptr++;
	}
}

function dcs_TAG(TagImage){
	var P ="http"+(window.location.protocol.indexOf('https:')==0?'s':'')+"://"+TagImage+"/dcs.gif?";
	for (N in DCS){P+=A( N, DCS[N]);}
	for (N in WT){P+=A( "WT."+N, WT[N]);}
	for (N in DCSext){P+=A( N, DCSext[N]);}
	saveReqArgs = P;
	dcs_createImage(P);
}

dcs_var();
dcs_TAG(TagPath);
//-->
</SCRIPT>
      <NOSCRIPT><IMG height=1 src="images/njs.gif" width=1 
      border=0 name=DCSIMG> </NOSCRIPT>                     
      <TABLE cellSpacing=0 cellPadding=0 width=758 border=0>
        <TBODY>
        <TR>
          <TD>
<!---- Main Content is inserted here ------>         



    <form name="form1" action="Order_c.asp" method="post">  
<%	strsql="Select * from Acclist where A_id= "& session("carno") &" ORDER BY Acclist.id DESC  "	
	
	set acres=conn.execute(strsql)				
	if not Acres.eof then 
					%>
		    	   

          <table width="600" cellspacing="0" cellpadding="0" border="1">
            <tr align="center"   > 
              <td colspan="7" height="20"> 
                       </td>
            <tr align="center"   > 
              <td colspan="7" height="40"> 
                <p><font size="2" color="#000000" face="Arial, Helvetica, sans-serif">樣板列表</font></p>
              </td>
            </tr>
            <tr align="center"  > 
              <td width="153" height="31"><font size="1" color="#000000" >型號</font></td>
              <td width="154" height="31"><font size="1" color="#000000" >數量</font></td>
              <td width="131" height="31"><font size="1" color="#000000" >價格</font></td>
              <td width="152" height="31"><font size="1" color="#000000" >總數量</font></td>
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
            <tr   align="center"> 
              <td width="153"> 
                <font size="1" color="#000000" > 
                <%
				set w=createobject("adodb.recordset")
           		w.CursorType = adOpenKeyset
                w.LockType=adLockOptimistic
                sql="SELECT  id FROM product  where model='"&trim(acres("P_model"))&"'"
                w.open sql,conn
                recount=w.recordcount
                if not w.eof then
               %>
                </font>
                <font  color="#000000"  color="#000000" ><a href="javascript:;"  onClick="openBrWindow('detail_c.asp?id=<%=trim(w("id"))%>')" title="detail"><%=trim(Acres("P_model")&"") %></a> 
                <input type="hidden" name="txtpm" value="<%= trim(Acres("P_model")&"") %>">
                </font></td>
              <td width="154"> 
                <font size="1" color="#000000" > 
                <input type="text" name="txtcount" value="<%=Acres("qty")%>" maxlength="5">
                </font>
              </td>
              <td width="131"> 
                <font size="1" color="#000000" > 
                <% if trim(acres("P_price"))=0 then %>
                Contact us   
                <%else%>
                <%= trim(acres("p_type")&"")&trim(Acres("P_price")&"")%>&nbsp;   
                <%end if%>
                <input type="hidden" name="txtpc" value="<%= trim(Acres("P_price")&"")%>">  
                </font> 
              </td>
              <td width="152">  
                <font size="1" color="#000000" >  
                <% if trim(acres("total"))=0 then %>
                Contact us            
                <%else%>
                <%= trim(acres("p_type")&"")&trim(Acres("total")&"")%>&nbsp;            
                <%end if%>
                <input type="hidden" name="txttot" value="<%= trim(Acres("total")&"") %>">           
                </font>  
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
              <td colspan="7" height="20"> 
                       </td>
 <tr align="center"   > 
              <td height="29"><font size="1" color="#000000" >公司名稱:</font></td>           
              <td height="14" colspan="6"> 
                <div align="left"> 
                  <font size="1" color="#000000" > 
                  <input type="text" name="cname" size="30" value="<%=cName%>" maxlength="33">
                  </font>
                </div>
              </td>
            </tr>
 <tr align="center"   > 
              <td height="29"><font size="1" color="#000000" >姓:</font></td>  
              <td height="14" colspan="6"> 
                <div align="left"> 
                  <font size="1" color="#000000" > 
                  <input type="text" name="fname" size="30" value="<%=FName%>" maxlength="33">
                  </font>
                </div>
              </td>
            </tr>
<tr align="center"   > 
              <td height="29"><font size="1" color="#000000" >名:</font></td>  
               <td colspan="6" height="14"> 
                <div align="left"> 
                  <font size="1" color="#000000" > 
                  <input type="text" name="lname" size="30" value="<%=LName%>" maxlength="33">
                  </font>
                </div>
              </td>
            </tr>
<tr align="center"   > 
              <td height="29"><font size="1" color="#000000" >電郵:</font></td>
              <td height="14" colspan="6"> 
                <div align="left"> 
                  <font size="1" color="#000000" > 
                  <input type="text" name="email" size="30" value="<%=email%>" maxlength="33">
                  </font>
                </div>
              </td>
            </tr>
            <tr align="center"   > 
              <td rowspan="3" height="29"><font size="1" color="#000000" >收貨地址</font></td>
              <td colspan="6" height="14"> 
                <div align="left"> 
                  <font size="1" color="#000000" > 
                  <input type="text" name="sd1" size="30" value="<%=sd1%>" maxlength="33">
                  </font>
                </div>
              </td>
            </tr>
            <tr align="center"   > 
              <td colspan="6" height="7"> 
                <div align="left">
                  <font size="1" color="#000000" >
                  <input type="text" name="sd2" size="30" value="<%=sd2%>" maxlength="33">
                  </font>
                </div>
              </td>
            </tr>
            <tr align="center"   > 
              <td colspan="6" height="8"> 
                <div align="left">
                  <font size="1">
                  <input type="text" name="sd3" size="30" value="<%=sd3%>" maxlength="33">
                  </font>
                </div>
              </td>
            </tr>
            <tr align="center"   > 
              <td height="7" colspan="7"> 
                <input type="button" value="繼續溜灠" onClick="dosubmit();" name="button">
                <input type="button" name="Button2" value="提交" onClick="dofin();"> 
                <input type="button" value="更改"  name="button2" onClick="dofresh();"> 
                <input type="button" name="Button" value="清空樣板籃" onClick="dodelect();"> 
              &nbsp;&nbsp;
              <a href="info.asp">Term and Condition</a></td>
            </tr>
          </table>
          <%else%>
            <TABLE cellSpacing=0 cellPadding=0 width=90% height=300 border=0>
             <tr>
               <td>
          <p align="center"><b><font size="3">你的樣板籃裹沒有東西,</font></b></p> 
		   <p align="center"><b><font size="3">請繼續 <a href="SearchResult_c.asp"> 搜索及查詢</a> !</font></b></p> 
		   </td>
		   </tr>
		   </table> 
 <%end if 
'set strsql=nothing
%>

<input type="hidden" name="action" value="">
</form>
 



            
<!---- Main Content is ended here ------>  

<!--#include file="menu_bottom_c.inc"-->
</BODY>
<%  
function showpageno(pageno) 
	if recount>10 then 
		lastpage=recount\10 
		if recount mod 10 >0 then 
			lastpage=lastpage+1 
		end if 
		if pageno>10 then 
		     response.write "<a href='order_c.asp?pageno=1'> The First Page </a>&nbsp;&nbsp;" 
			response.write "<a href='order_c.asp?pageno="&(pageno-9-(pageno  mod 10) )&"'>Previous 10</a>&nbsp;&nbsp;" 
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
				response.write "<a href='order_c.asp?Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;" 
			end if	 
		next 
		if (pageno\10)<(lastpage\10) then 
		        response.write "<a href='order_c.asp?Pageno="&(pageno+11-(pageno mod 10)) &"'>Next 10</a>&nbsp;&nbsp;" 
			   response.write "<a href='order_c.asp?Pageno="& (lastpage) &"'> The Last Page</a>&nbsp;&nbsp;" 
		end if 
		 
 end if 
end function 
%>

