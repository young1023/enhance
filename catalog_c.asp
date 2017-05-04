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
            
<table border="0" width="98%" bordercolor="#FF0000" bordercolorlight="#FF00FF" bordercolordark="#C0C0C0" cellspacing="8" cellpadding="8" height="250"><tr>
    <td width="25%"><font color="#000000">&nbsp;</font></td>
    <td width="25%"><font color="#FF0000">&nbsp;</font></td>
    <td width="25%"><font color="#FF0000">&nbsp;</font></td>
    <td width="25%"><font color="#FF0000">&nbsp;</font></td>
  </tr>
 <tr>
    <td width="25%"><font color="#000000">2004年 目錄</font></td>
    <td width="25%"><font color="#FF0000"><a href="files/uploadcat/2004cat.zip">下載2003Cat.zip目錄</a></font>&nbsp;</td>
    <td width="25%"><font color="#FF0000"><a href="files/uploadcat/2004cat.dnl">下載2003Cat.dnl目錄</a></font>&nbsp;</td>
    <td width="25%"><font color="#FF0000"><a href="files/uploadcat/2004cat.htm">觀看網上目錄</a></font></td>
  </tr>
 <tr>
    <td width="25%"><font color="#000000">2003年 目錄</font></td>
    <td width="25%"><font color="#FF0000"><a href="files/uploadcat/2003cat.zip">下載2003Cat.zip目錄</a></font>&nbsp;</td>
    <td width="25%"><font color="#FF0000"><a href="files/uploadcat/2003cat.dnl">下載2003Cat.dnl目錄</a></font>&nbsp;</td>
    <td width="25%"><font color="#FF0000"><a href="files/uploadcat/2003cat.htm">觀看網上目錄</a></font></td>
  </tr>
   <tr>
    <td width="25%"><font color="#000000">2002年 目錄</font></td>
    <td width="25%"><font color="#FF0000"><a href="files/uploadcat/2002cat.zip">下載2002Cat.zip目錄</a></font>&nbsp;</td>
    <td width="25%"><font color="#FF0000"><a href="files/uploadcat/2002cat.dnl">下載2002Cat.dnl目錄</a></font>&nbsp;</td>
    <td width="25%"><font color="#FF0000"><a href="files/uploadcat/2002cat.htm">觀看網上目錄</a></font></td>
  </tr>
  <tr>
    <td width="25%"><font color="#000000">2001年 目錄</font></td>
    <td width="25%"><font color="#FF0000"><a href="files/uploadcat/2001cat.zip">下載2001Cat.zip目錄</a></font>&nbsp;</td>
    <td width="25%"><font color="#FF0000"><a href="files/uploadcat/2001cat.dnl">下載2001Cat.dnl目錄</a></font>&nbsp;</td>
    <td width="25%"><font color="#FF0000"><a href="files/uploadcat/2001cat.htm">觀看網上目錄</a></font></td>
  </tr>
 <tr>
    <td width="25%"><font color="#000000">2000年 目錄</font></td>
    <td width="25%"><font color="#FF0000"><a href="files/uploadcat/2000cat.zip">下載2000Cat.zip目錄</a></font>&nbsp;</td>
    <td width="25%"><font color="#FF0000"><a href="files/uploadcat/2000cat.dnl">下載2000Cat.dnl目錄</a></font>&nbsp;</td>
    <td width="25%"><font color="#FF0000"><a href="files/uploadcat/2000cat.htm">觀看網上目錄</a></font></td>
  </tr>
<tr>
<tr>
    <td colspan="4"><font color="#000000">請在此下載Digital WebBook DNL reader 收看DNL檔案，閣下可以點選頁面或型號觀看詳細資料，也可以以電郵方式轉寄其他同事，<br>這目錄特定為你們設計歡迎分享。</font></td>
    </tr>
<tr>
    <td colspan="4"><a href="http://www.digitalwebbook.com/reader/"><img src="http://www.digitalwebbook.com/images/getdnlreader.gif">
      <font color="#FF0000">下載 DigitalWebBook DNL Reader  </a></font></td>
  </tr>
<tr>
    <td colspan="4"><a href="http://www.3dbrochure.com/default.asp?aid=DTAWINSOWRGBUX">
<img src="/images/birdshape.gif" border="0" alt="">
</a>
  </tr>
</table>

<!---- Main Content is ended here ------>  

</TD>
<!--#include file="menu_bottom_c.inc"-->
</BODY>