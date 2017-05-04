<!--#include file="included/conne.inc" -->
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
              
<!-- content is inserted here ---->              
                          
           
  <TABLE border=0 cellPadding=0 cellSpacing=0 height=100% width=99%>             
    <TBODY>              
    <TR>             
                
        
      <TD vAlign=top>             
        <table width="99%" border="0" cellpadding="1" cellspacing="0">             
          <tr>             
            <td><img src="images/back.gif" width="306" height="37"></td>             
          </tr>             
        </table>             
                  
        <table width="100%" border="0" cellpadding="1" cellspacing="1" height="353">             
          <tr>              
            <td bgcolor="#000000">              
              <table width="100%" border="0" cellpadding=0 cellspacing="0" bgcolor="#FFFFFF">             
                <tr>             
                  <td height="353" valign="top">             
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">             
                      <tr>              
                        <td height="28">&nbsp;<font color="#FF9900">●</font>&nbsp;&nbsp;<b>公司新聞</b></td>             
                      </tr>             
                      <tr>              
                        <td>              
                          <%             
				  subsql="select * from message order by typeid desc"             
				  set subrs=conn.execute(subsql)             
				  i=1             
				  do while not subrs.eof             
                    'if subrs.eof or i>10 then exit do             
                    if subflage then             
                      submycolor="#ffffff"             
                    else             
                      submycolor="#efefef"             
                    end if             
					  %>             
                          <table width="100%" border="0" cellspacing="2" cellpadding="0">             
                            <tr bgcolor=<%=submycolor%>>             
                              <td height="23" width="80%" onmouseover="javascript:style.background='#FFCC33'" onmouseout="javascript:style.background='<%=submycolor%>'">             
							  &nbsp;&nbsp;&nbsp;&nbsp;<span onclick="javascript:window.open('<%=subrs("http")%>','user','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,top=80,left=160,width=800,height=600');" style="cursor:hand"><included src="included/row.gif" width="10" height="8">&nbsp;&nbsp;<%=replace(trim(server.htmlencode(subrs("title"))),vbcrlf,"<br>")%></span>             
                              </td>             
                              <td><font color=blue><%=subrs("createtime")%></font></td>             
                            </tr>             
                          </table>             
                          <%             
				      subrs.movenext             
					  subflage=not subflage             
					  i=i+1             
				  loop             
				  subrs.close             
				  set subrs=nothing             
					  %>             
                        </td>             
                      </tr>             
                      <tr>             
                        <td height="8"></td>             
                      </tr>             
                    </table>             
                    <%             
             
conn.close             
set conn=nothing             
%>             
                  </td>             
                </tr>             
              </table>             
                  </td>             
                </tr>             
              </table>       
              
            </td>            
                </tr>            
              </table>      
<!-- content is ended here ---->              
              
</TD>              
              
              
              
              
              
              
              
              
              
<!--#include file="menu_bottom_c.inc"-->              
</BODY>              
