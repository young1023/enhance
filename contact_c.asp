<!--#include file="database.inc"-->

<%
act_bt=request("txtmail")
straction=request("action_button")
email=request("email")

if straction="save" then         
         fn=replace(request("name"),"'","''")
         title=replace(request("title") ,"'","''")
         em=replace(request("email"),"'","''")      
		 cn=replace(request("company"),"'","''")
         address=replace(request("address"),"'","''")
		 fax=replace(request("fax"),"'","''")
		 phone=replace(request("phone"),"'","''")
		 product=replace(request("Product"),"'","''")
		 session("product")=product
        comment=replace(request("comment"),"'","''")				 
		if trim((comment)&"")="" then
		    comment=null
		end if
		dat=date()
	sqlcmd="select*from member where email='"&em&"'"
    'response.write sqlcmd
	set rs=conn.execute(sqlcmd) 
	
	if not rs.eof then
        strsql1="update member set firstname= '"& fn & "' ," _
			          &"title='"& title &"'," _
			          &"phone='"& phone &"'," _
			          &"fax='"& fax & "'," _
			          &"CompanyName='"& cn & "', " _
					  &"comment='"& comment & "' ," _
					   &"create_date=#"& dat & "# ," _
					   &"address='"& address & "' " _
					   &"WHere email='"& em&"'  "
	     	'response.write strsql1
			     conn.execute strsql1
			     set strsql1=nothing 
   end if
   if rs.eof then 
		sQLcmd= "insert into member( title,"
		sqlcmd=sqlcmd&"fax,phone,Firstname,"
		sqlcmd=sqlcmd&"address,Email,"
		sqlcmd=sqlcmd&"CompanyName,comment,create_date) "
		sqlcmd=sqlcmd&"values ('"& title &"','"& fax 
		sqlcmd=sqlcmd&"','"& phone &"','"& fn
		sqlcmd=sqlcmd &"','"& address &"','"& em 
		sqlcmd=sqlcmd&"','"& cn &"','"& comment &"',#"& dat &"#)"
		'response.write sqlcmd
	    conn.Execute sqlcmd
	end if 
	sqlcmd="select*from member where email='"&em&"'"
	set rs=conn.execute(sqlcmd)
	'response.write rs.eof&"ppppp"
   if  not rs.eof then 
     session("user_id")=rs("id")
     response.redirect "enquiry_result.asp"
    ' response.redirect "sendemail.asp?email="&server.Urlencode(email)
	end if  
end if
if straction=""then
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
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
function dosubmit()
{
	var errmsg;
	var strmail;
	errmsg="";
	strmail="";
	 strpwd=document.updata.Product.value;
	if(strpwd=="")
	errmsg=errmsg+"Please Input Product& Qty.\n";	
	if(document.updata.name.value=="")
	errmsg=errmsg+"Please Input  name.\n";
	if(document.updata.title.value=="")
	errmsg=errmsg+"Please Input  title.\n";	
    if(document.updata.company.value=="")
	errmsg=errmsg+"Please Input company Name.\n";
	if(document.updata.address.value=="" )
	errmsg=errmsg+"Please Input address.\n";
	strmail=document.updata.email.value;
	if(strmail=="" )
		errmsg=errmsg+"Please Input e-mail address.\n";
	else
		if(strmail.search("@")==-1)
			errmsg=errmsg+"Please Input a right e-mail address.\n";
	//phone=document.updata.phone.value
	//if(phone=="" )
	//	errmsg=errmsg+"Please Input phone.\n";
	//else
	//	if(isNaN(phone))
	//		errmsg=errmsg+"Please Input a right phone.\n";
	//fax=document.updata.fax.value
	//if(fax=="" )
	//	errmsg=errmsg+"Please Input FAX.\n";
	//else
	//	if(isNaN(fax))
	//		errmsg=errmsg+"Please Input a right FAX.\n";  
	
	if(errmsg!="")
	{
		alert(errmsg);
	}
	else
	{
		document.updata.action_button.value="save";
		document.updata.submit();
	}
}
</script>

<!--#include file="menu_bar_c.js"-->

<!--END menuItemData.asp -->
</HEAD>
<BODY bgColor=#ffffff leftMargin=2 topMargin=2 marginwidth="2" marginheight="2" text="#FF0000" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<font face="Arial" color="#000000" size="1">
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
      <TABLE cellSpacing=0 cellPadding=0 width=758 border=0>
        <TBODY>
        <TR>
          <TD>
<!---- Main Content is inserted here ------>         
            
 
  <form name="updata" mothed="post" action="contact_c.asp">
        <div align="center">
        <div align="center"><center><p><font face="Arial" color="#000000" size="1">If you would like more information about our company <br> 
          please fill out the form below. <br> 
          We will send you a CD-ROM which included all of our products.</font></p> 
          </center></div><div align="center"><center><table border="1" width="588">
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" color="#000000" size="1">Product &amp; Qty :<br> 
                e.g. Model No.,<br> 
              Metal Ball Pen 
              </font></td>
			  
 <%
  if session("enquiryno")<>"" then
 	strsql="Select * from Acclist where A_id= "& session("enquiryno") &" ORDER BY Acclist.id DESC  "	
	
	set acres=conn.execute(strsql)				
	if not Acres.eof then 
	Do while (not acres.eof)      
         strtemp=strtemp&chr(10)&trim(Acres("P_model")&"")  &space(3)& Acres("qty")   
         acres.movenext					
    loop  
    end if
%>
      <td width="448"><font size="1"><textarea rows="5" name="Product" cols="36"><%=strtemp%></textarea><font face="Arial" color="#000000">(required)</font></font></td>
<% else %>
                  <td width="448"><font size="1">
                    <textarea rows="5" name="Product" cols="36"></textarea>
                    <font face="Arial" color="#000000">(required)</font></font></td>
<% end if%>			  		
            </tr>
            <tr>
              <td width="128" align="right" valign="top">¡@</td>
              <td width="448"><font face="Arial" color="#000000" size="1"><input type="checkbox"
              name="Send_Product_Literature" value="ON">  Send Product Literature</font></td> 
            </tr>
            <tr>
              <td width="128" align="right" valign="top"><font size="1">&nbsp;</font></td>
              <td valign="top" width="448"><font face="Arial" color="#000000" size="1"><input
              type="checkbox" name="Send_CD" value="ON">  Send CD</font></td> 
            </tr>
            <tr>
              <td width="128" align="right" valign="top"><font size="1">&nbsp;</font></td>
              <td width="448"><font face="Arial" color="#000000" size="1"><input type="checkbox"
              name="Have_a_Salesperson_Contact_Me" value="ON">  Have a Salesperson Contact Me</font></td> 
            </tr>
			<%
	 if session("user_id")<>"" then
			 sql="select lastname,email,title,companyname,address,phone,fax from member where id="& session("user_id") &" "
             set w=conn.execute(sql)
             if not w.eof then 
			   name=w("lastname")
			   title=w("title")
			   company=w("companyname")
			   address=w("address")
			   email=w("email")
			   phone=w("phone")
			   fax=w("fax")
			  end if
	  end if
			%>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" color="#000000" size="1">Name 
              :</font></td>
              <td width="448"><font face="Arial" color="#000000" size="1">
                    <input type="text" size="32"
              name="name" value="<%=name%>">
                    (required)
                    </font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" color="#000000" size="1">Title 
              :</font></td>
              <td width="448"><font face="Arial" color="#000000" size="1">
                    <input type="text" size="32"
              name="title" value="<%=title%>">
                    (required)
                    </font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" color="#000000" size="1">Company 
              :</font></td>
              <td width="448"><font face="Arial" color="#000000" size="1">
                    <input type="text" size="32"
              name="company" value="<%=company%>">
                    (required)
                    </font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" color="#000000" size="1">Address 
              :</font></td>
                  <td width="448"><font face="Arial" color="#000000" size="1">
                    <textarea name="address"
              rows="5" cols="31"><%=address%></textarea>
                    (required)
                    </font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" color="#000000" size="1">E-mail 
              :</font></td>
              <td width="448"><font face="Arial" color="#000000" size="1">
                    <input type="text" size="25"
              name="email" value="<%=email%>">
                    (required)
                    </font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" color="#000000" size="1">Phone 
              :</font></td>
              <td width="448"><font face="Arial" color="#000000" size="1">
                    <input type="text" size="25"
              name="phone" value="<%=phone%>">
                    (required)
                    </font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" color="#000000" size="1">FAX 
              :</font></td>
              <td width="448"><font face="Arial" color="#000000" size="1">
                    <input type="text" size="25"
              name="fax" value="<%=fax%>">
                    (required)
                    </font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" color="#000000" size="1">Comment &amp; Message :</font></td> 
              <td width="448"><font face="Arial" color="#000000" size="1"><textarea rows="5"
              name="Comment" cols="31"></textarea></font></td>
            </tr>
            <tr>
              <td width="128" align="right" valign="top"><font size="1">&nbsp;</font></td>
              <td valign="bottom" width="448"><font face="Arial" color="#000000" size="1">
			  <input type="button" value="Submit Request" onClick="javascript:dosubmit();"> 
			  <input type="reset" value="Reset Form">  
			  <input type="hidden" name="action_button" value=""></font></td> 
            </tr>
          </table></center></div></div>
		  </form>
    </td>
  </tr>
  <tr>
    <td  width="644" valign="top"><font size="1">&nbsp;</font></td>
  </tr>
</table>
<div align="center">
  <center>
    <table border="0" cellpadding="0" cellspacing="0" width="593">
      <tr>
        <td width="250"></td>
        <td width="343"></td>
    </tr>
    <tr>
        <td width="250"><u><font face="Arial" color="#000000" size="1">Head  
          Office</font></u> </td>
        <td width="343"></td>
    </tr>
    <tr>
        <td width="250"><font face="Arial" color="#000000" size="1">1/F, Air Goal Cargo Bldg.,<br>330 Kwun Tong Road, <br>Kowloon, Hong Kong<br> 
    Tel: (852) 2798 6263 (12 lines)<br> 
    Fax: (852) 2751 6659 (2 lines)</font></td> 
        <td width="343" valign="top"><font face="Arial" color="#000000" size="1">Flat  
          C, 1/F., Kam Mong Bldg.<br> 
        37-39 Fa Yuen street, Kln. H.K.<br> 
    Tel: (852) 2771 1333, 2780 1029<br> 
    Fax: (852) 2770 0351</font></td> 
    </tr>
    <tr>
        <td colspan="2"> 
          <p align="center"><font size="1"><font face="Arial" color="#000000">
    E-mail: <a href="mailto:info@jaguarpen.com">info@jaguarpen.com</a> 
          </font>


<!---- Main Content is ended here ------>  

          </font>  

</TD>
<!--#include file="menu_bottom_c.inc"-->
</BODY>
<% end if%>