<!--#include file="database.inc"-->
<% 
if session(("SECURITY_LEVEL"))<>"1" then
				response.redirect "default.asp"
end if
strid=trim(request("A_id"))
'response.write strid
if strid="" then
	response.redirect "admin_order.asp"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="imagetoolbar" content="no">
<TITLE>Enhance Techologies Limited</TITLE>
<LINK title=style1 
href="images/basic.css" type=text/css rel=stylesheet>
<SCRIPT language=JavaScript src="images/js_dynMenuFunctions.js">
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
</script>
<!--#include file="menu_bar.inc"-->

<!--END menuItemData.asp -->
</HEAD>
<BODY bgColor=#ffffff leftMargin=2 topMargin=2 marginwidth="2" marginheight="2" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<SCRIPT language=javascript>
function openWin(arg_url, arg_name, arg_width, arg_height, arg_scroll, arg_resize) {
    var win = window.open(arg_url,arg_name,'scrollbars=' + arg_scroll + ',resizable=' + arg_resize + ',width=' + arg_width +  ',height=' + arg_height);
  }
//-->
</SCRIPT>
<!--#include file="menu_nav.inc"-->

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

<!-- content is inserted here ---->
            

<table border="0" cellspacing="0" cellpadding="0" width="800" height="200">
  <tr>
       <td valign="top" align="center" width="678" > 
<br>          
            <p><a href="admin_Order.asp">回前頁</a> / 詢價資料</p>
  <p> 
    <% 
		strsql="Select * from Acclist  where A_Id= "& strid &"  "
		'response.write strsql
		set Acres1=Conn.Execute(strsql)
		IsExist=false
  		if not Acres1.eof then
			IsExist=true
		end if
	
  if IsExist  then
  %>
  </p>
</div>
<div align="center">
            <table width="90%" border="0" cellspacing="1" bgcolor="#008080" >
              <tr> 
                <td width="42%" bgcolor="#FFFFFF"> 
                  <div align="center">Model</div>
                </td>
                <td width="12%" bgcolor="#FFFFFF"> 
                  <div align="center">Qty</div>
                </td>
              </tr>
              <%
	
	
	 Do while not AcRes1.eof 		
             %>
              <tr> 
                <td align="center" width="42%" bgcolor="#FFFFFF"><%= trim(Acres1("p_model") &"")%>&nbsp;</td>
                <td align="center" width="12%" bgcolor="#FFFFFF"> <%= trim(Acres1("qty") &"")%>&nbsp;</td>
              </tr>
              <% 
			  sd=trim(acres1("saddress"))
	Acres1.movenext
	  loop 
         on error  resume next
         if sd<>"" then
			sd=split(sd,"^")
			i=Ubound(sd)
			if i=0 then
		      sd1=trim(sd(0))
			  'sd2=""
		    elseif i=1 then
		      sd1=trim(sd(0))
			  sd2=trim(sd(1))
			  'sd3=""
			elseif i=2 then
               sd1=trim(sd(0))
			   sd2=trim(sd(1))
			   sd3=trim(sd(2))
            end if
	        %>
              <%end if%>
            </table>
   </div>
<div align="center">
  <% End If 
  if not IsExist then
  %>
 <p align="center"><font color="#FF3366">No Product List</p>
<% End If %>
          
    </div>
  </tr>
</table>

<!-- content is ended here ---->

</TD>









<!--#include file="menu_bottom.inc"-->
</BODY>
