<!--#include file="database.inc"-->
<%
'on error resume next
response.expires=0
if session("carno")="" then
  response.redirect "searchresult.asp"
end if
stremail=trim(request("txtemail"))
strpassword=trim(request("txtpassword"))
FName=session("FName")
LName=session("LName")
cName=session("cName")
Email=Session("email")

'Response.end
if stremail<>"" and strpassword<>"" then
  strsql="Select * from member where email='" & stremail & "' and mypassword='"& strpassword & "'"
  'response.write strsql
  set acres=conn.execute(strsql)
  if  not acres.eof then
	session("user_id")=trim(acres("id")&"")
	session("SECURITY_LEVEL") = trim(acres("level")&"")
	session("em")=trim(acres("email")&"")
	session("id")=trim(acres("id")&"")
	session("admin")=trim(acres("level")&"")
	saddress=trim(acres("saddress"))			
  else
	strerr="Member login failure, "
  end if
end if	
'If session("user_id")<>"" and session("carno")<>"" then
If session("carno")<>"" then

  strsql="Select * from Acclist  where A_Id="& session("carno") &"   "
  'response.write strsql
  set Acres1=Conn.Execute(strsql)
  if not Acres1.eof then
    strid=session("carno")
	strsql="Select Max(order_no) from AccMain" 
	set AcRes=Conn.Execute(strsql)
	'response.write strsql
	if not Acres.EOF then
	  if trim(acres("expr1000"))<>""then			    
	    orderno=trim(acres("expr1000"))+1
	    'response.write orderno
        'response.end
	  else 
	    orderno=1
	  end if
    end if
    set AcRes=nothing
    set strsql=nothing 

If session("user_id")="" then

   ':: Add Customer Detail into database

     strsql1="insert into Member" _
     & "(Email,CompanyName,FirstName,LastName) values " _
					  &" ('"&Email&"','" _
                      &"  "&cName&"','" _
					  &"  "&FName&"','" _
					  &" "&LName&"')"
					  'response.write strsql1
                      'response.end
				  Conn.Execute strsql1
                 set strsql1=nothing 

     
   ':: Get the Customer ID
      strsql2="Select top 1 * from Member order by id desc" 
	set AcRes2=Conn.Execute(strsql2)

	  CustomerID= AcRes2("ID")

      set AcRes2=nothing
      set strsql2=nothing
  
	'response.write CustomerID
    'response.end	

    strsql="insert into accmain(U_Id,U_No,Order_NO,AccDate) Values(" _
     &" "&CustomerID&", " _
     &" "&session("carno")&" ," _
     &" "& trim(orderno) & " ," _
     &"'"&date& " " & time&"')"
    'response.write strsql
    'response.end

    	End if

If session("user_id")<>"" then


 strsql="insert into accmain(U_Id,U_No,Order_NO,AccDate) Values(" _
     &" "&session("user_id")&", " _
     &" "&session("carno")&" ," _
     &" "& trim(orderno) & " ," _
     &"'"&date& " " & time&"')"
    'response.write strsql
   ' response.end

End if

    Conn.Execute strsql
    set AcRes=nothing
    set strsql=nothing
    strcarno=session("carno")
  else
    strid=session("carno")
  end if
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

<!--#include file="menu_bar_c.js"-->

<!--END menuItemData.asp -->
</HEAD>
<BODY bgColor=#ffffff leftMargin=2 topMargin=2 marginwidth="2" marginheight="2" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<font face="Arial" size="1">
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
</font>        
      <TABLE cellSpacing=0 cellPadding=0 width=758 border=1 bordercolor="#FF0000">
        <TBODY>
        <TR>
          <TD>
<!---- Main Content is inserted here ------>         
            
  
     <%'if session("user_id")<>""and session("carno")<>""then 
  if session("carno")<>"" then 
    
      session("carno")=""
		 strsql="Select * from Acclist  where A_Id="& strid &"   "
		 'response.write strsql
		 set Acres1=Conn.Execute(strsql)
		 IsExist=false
  		 if not Acres1.eof then
		  IsExist=true
         end if	
         if IsExist  then
     %>
<table border="0" width="80%">
  <tr>
    <td width="20%"><img border="0" src="images/visa.jpg" width="80" height="53"></td>
    <td width="20%"><img border="0" src="images/visadelta.jpg" width="80" height="62"></td>
    <td width="20%"><img border="0" src="images/visa-electron.jpg" width="80" height="51"></td>
    <td width="20%"><img border="0" src="images/mastercard.jpg" width="80" height="48"></td>
  </tr>
  <tr>
    <td width="20%"></td>
    <td width="20%"></td>
    <td width="20%"></td>
    <td width="20%"></td>
  </tr>
</table>
      <p><font size="2" color="#000000">Ordering Information</font></p>  
      <p><font size="2" color="#000000">Order Number:<%=orderno%></font></p> 
      <p><font size="2" color="#000000">"Refunds will be given at the discretion of the Company Management"</font></p> 
      <form name=w action="https://select.worldpay.com/wcc/purchase" method=POST>
          
        <table width="88%" border="1" cellspacing="1" cellpadding="1" bordercolor="#FF00FF" bgcolor="#FFFF99" >
          <tr  > 
            <td colspan="4" height="20" align="center" bgcolor="#FFFF00"><font size="2" color="#000000">網上結算</font></td>
          </tr>
          <tr  > 
            <td width="26%" align="center"> <font size="1" color="#000000">型號</font> 
            </td>
            <td width="27%" align="center"> <font size="1" color="#000000">數量</font> 
            </td>
            <td width="23%" align="center"> <font size="1" color="#000000">價格</font> 
            </td>
            <td width="24%" align="center"> <font size="1" color="#000000">總數</font> 
            </td>
          </tr>
          <%	
              Do while not AcRes1.eof		
             %>
          <tr  > 
            <td align="center" width="26%"><font size="1" color="#000000"><%= trim(Acres1("p_model") &"")%>&nbsp;</font></td>
            <td align="center" width="27%"> <font size="1" color="#000000"><%= trim(Acres1("qty") &"")%>&nbsp;</font></td>
            <td align="center" width="23%"><font size="1" color="#000000"> 
              <%if trim(acres1("P_price")&"") =0 then
                    p_price=0
                    response.write "<font size=2>Contact Us </font>"
                  else
                    p_price=trim(acres1("p_price"))
				  %>
              </font><font color="#000000"> 
              <%=trim(acres1("p_type")&"")&trim(Acres1("p_price")&"")%> 
              <%end if%>
              </font></td>
            <td align="center" width="24%"><font size="1" color="#000000"> 
              <%if trim(acres1("total")&"")=0 or trim(acres1("total")&"")="" then %>
              Contact Us      
              <%else 
                    yourcount=acres1("total")+yourcount
                    whatsale=trim(Acres1("p_model"))+","+trim(Acres1("qty")) +","+p_price+","+trim(acres1("total"))+";"+whatsale
				  %>
              <%=trim(acres1("p_type")&"")&trim(Acres1("total")&"")%>&nbsp;       
              <%end if%>
              </font></td>
          </tr>
          <%Acres1.movenext
                whatsale=replace(whatsale,";","<br>")
               loop
			  %>
          <tr  > 
            <td align="center" width="26%"> 
              <font size="1" color="#000000"> 
              <input  type=hidden name="instId" value="38019">
              <input type=hidden name="cartId" value="35670">
              <input type=hidden name="testMode" value="0">
              <input type=hidden name="amount" value="<%=yourcount%>">
              <input type=hidden name="desc" value="<%=trim(whatsale)%>">
              <input type=hidden name="currency" value="USD">
              <input type=hidden name="account" value="5609906">
              &nbsp;</font> </td>
            <td align="center" width="27%"><font size="1" color="#000000">&nbsp;</font> </td>
            <td align="center" width="23%"><font size="1" color="#000000">&nbsp;</font></td>
            <td align="center" width="24%"><font size="1" color="#000000">USD <%=yourcount%>       
              </font></td>
          </tr>
          <tr  > 
            <td align="center" width="26%" height="16"><font size="1" color="#000000">&nbsp;</font> 
              
            </td>
            <td align="center" colspan="2" height="16"><font size="1" color="#000000">&nbsp;</font></td>
            <td align="center" width="24%" height="16"><font size="1" color="#000000"> 
              <input type=submit value="Buy It" name="submit">
              </font></td>
          </tr>
          <tr> 
            <td align="center" colspan="4"> 
              <font size="1" color="#000000"> 
              <% session("carno")="" %>
              <a href="SearchResult_c.asp">Go       
              Back</a> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <a href="info.asp" target="_blank">Term and Condition</a></font> </td>
          </tr>
          <tr> 
            <td align="center" colspan="4"><font size="1" color="#000000">&nbsp;</font></td>
          </tr>
        </table>
        </form>
       <%end if
         if IsExist=false then
           session("carno")=""
        %>
           <p align="center"><font size="1" color="#000000">This order was delivered,</font></p>      
          <p align="center"><font size="1" color="#000000"> please <a href="SearchResult_c.asp">search and order</a> again !</font></p> 
        <%End If 
	end if
'end if
 %>

  


<!---- Main Content is ended here ------>  

<!--#include file="menu_bottom_c.inc"-->
</table>
</BODY>
