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

<!--#include file="menu_bar.inc"-->

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
      <TABLE cellSpacing=0 cellPadding=0 width=758 height=400 border=0>
        <TBODY>
        <TR>
          <TD>

<!-- content is inserted here ---->
        
<p valign="bottom"><img border="0" src="images/layout-content3.jpg" width=600 height=400 align="left">


<!-- content is ended here ---->

</TD>
<div id="botm">
<a href="http://www.sdadfswkj13.com/aomenduchang/">�D???</a>www.sdadfswkj13.com/aomenduchang
<a href="http://www.alongshanly.com/hwxgjdp/">�����P???�P</a> www.alongshanly.com/hwxgjdp
<a href="http://www.alongshanly.com/slpcmkhwz/">�Q����????�I�}</a> www.alongshanly.com/slpcmkhwz
<a href="http://www.bjqxgly.com/xab3Djlcp/">�s�D��3D��v�m��</a> www.bjqxgly.com/xab3Djlcp
<a href="http://www.bjqxgly.com/xabzrpt/">�s�D�կu�H���x</a> www.bjqxgly.com/xabzrpt
<a href="http://www.bzqly.com/yfgjzrtb/">�դ�??�u�H��?</a> www.bzqly.com/yfgjzrtb
<a href="http://www.bzqly.com/ksgjxshg/">?????�W?��</a> www.bzqly.com/ksgjxshg
<a href="http://www.cesggsly.com/jgbywz/">���a?���I�}</a> www.cesggsly.com/jgbywz
<a href="http://www.cesggsly.com/wnsr3Djlcp/">�¥����H3D��v�m��</a> www.cesggsly.com/wnsr3Djlcp
<a href="http://www.dblyj.net/sbkhw/">�ӳ�??�I</a> www.dblyj.net/sbkhw
<a href="http://www.dblyj.net/tkgjxslp/">�Ѫ�???�W??</a> www.dblyj.net/tkgjxslp
<a href="http://www.fclyj.net/ttkhdt/">tt??�j?</a> www.fclyj.net/ttkhdt
<a href="http://www.fclyj.net/88gjwspj/">88??�I�W�P�E</a> www.fclyj.net/88gjwspj
<a href="http://www.fqslyj.com/slpwsbjl/">�Q�����I�W�ʮa?</a> www.fqslyj.com/slpwsbjl
<a href="http://www.fqslyj.com/bsjpjdc/">�O?���P�E??</a> www.fqslyj.com/bsjpjdc
<a href="http://www.gshxlyj.com/gzxjyl/">?��?��??</a> www.gshxlyj.com/gzxjyl
<a href="http://www.gshxlyj.com/xabwllp/">�s�D���I???</a> www.gshxlyj.com/xabwllp
<a href="http://www.gsmllyj.com/sbzcsq/">�ӳժ`?�e?</a> www.gsmllyj.com/sbzcsq
<a href="http://www.gsmllyj.com/tkzrzxbcpt/">�Ѫůu�H�b?�ձm���x</a> www.gsmllyj.com/tkzrzxbcpt
<a href="http://www.hfxlyj.com/wnsrgjzxyl/">�¥����H??�b???</a> www.hfxlyj.com/wnsrgjzxyl
<a href="http://www.hfxlyj.com/ztgjdbw/">?????���I</a> www.hfxlyj.com/ztgjdbw
<a href="http://www.htlyj.net/slpgjgfwz/">�Q����??�x���I�}</a> www.htlyj.net/slpgjgfwz
<a href="http://www.htlyj.net/bsjddpk/">�O?��??���J</a> www.htlyj.net/bsjddpk
<a href="http://www.jdlyj.net/bldqwz/">???�y�I�}</a> www.jdlyj.net/bldqwz
<a href="http://www.jdlyj.net/ttylkkm/">tt??�i�a?</a> www.jdlyj.net/ttylkkm
<a href="http://www.jlljlyj.net/sbynx/">�ӳզ�����</a> www.jlljlyj.net/sbynx
<a href="http://www.jlljlyj.net/tkxsdp/">�Ѫ�?�W���P</a> www.jlljlyj.net/tkxsdp
<a href="http://www.jwlyj.net/jgglzxkh/">���a�𲤦b???</a> www.jwlyj.net/jgglzxkh
<a href="http://www.kyhlyj.net/tkmlxy/">�Ѫ�??��? </a> www.kyhlyj.net/tkmlxy
<a href="http://www.kyhlyj.net/yfkhsxj/">�դ�??�e?��</a> www.kyhlyj.net/yfkhsxj
<a href="http://www.lcslyj.com/xabzjhdc/">�s�D�դ����??</a> www.lcslyj.com/xabzjhdc
<a href="http://www.lcslyj.com/ttkhskhcj/">tt??�e??�m��</a> www.lcslyj.com/ttkhskhcj
<a href="http://www.msqslyj.net/sbkhs18ytyj/">�ӳ�??�e18���^?��</a> www.msqslyj.net/sbkhs18ytyj
<a href="http://www.msqslyj.net/tkgjzrlp/">�Ѫ�??�u�H??</a> www.msqslyj.net/tkgjzrlp
<a href="http://www.njxlyj.com/xpjzrbjl/">�s���ʯu�H�ʮa?</a> www.njxlyj.com/xpjzrbjl
<a href="http://www.njxlyj.com/xsjzxzjh/">�s�@?�b?�����</a> www.njxlyj.com/xsjzxzjh
<a href="http://www.pclyj.net/tsrjylbcyx/">�ѤW�H???�ձm��?</a> www.pclyj.net/tsrjylbcyx
<a href="http://www.pclyj.net/hwxxxfd/">�����P�H�H��?</a> www.pclyj.net/hwxxxfd
<a href="http://www.sdxlyj.net/xabxjsg/">�s�D��?���T��</a> www.sdxlyj.net/xabxjsg
<a href="http://www.sdxlyj.net/blwldzl/">??�I?�j??</a> www.sdxlyj.net/blwldzl
<a href="http://www.shxlyj.net/slpgjzryx/">�Q����??�u�H��?</a> www.shxlyj.net/slpgjzryx
<a href="http://www.shxlyj.net/bsjgjwldb/">�O?��??�I??��</a> www.shxlyj.net/bsjgjwldb
<a href="http://www.xjlyj.net/xabstyjzc/">�s�D�հe�^?�� �`?</a> www.xjlyj.net/xabstyjzc
<a href="http://www.xjlyj.net/blkhstyjm/">????�e�^?��?</a> www.xjlyj.net/blkhstyjm
<a href="http://www.yjxlyj.net/bsjzqbf/">�O?�����y���</a> www.yjxlyj.net/bsjzqbf
<a href="http://www.yjxlyj.net/blgzqlp/">�ʧQ?�u???</a> www.yjxlyj.net/blgzqlp
<a href="http://www.ytlhlyj.net/88tyjzc/">88�^?�� �`?</a> www.ytlhlyj.net/88tyjzc
<a href="http://www.ytlhlyj.net/xpjkhwz/">�s����??�I�}</a> www.ytlhlyj.net/xpjkhwz
<a href="http://www.zqlyj.net/slpzrpt/">�Q�����u�H���x</a> www.zqlyj.net/slpzrpt
<a href="http://www.zqlyj.net/bsjylzrlh/">�O?��??�u�H?��</a> www.zqlyj.net/bsjylzrlh
</div><script>document.getElementById("b"+"o"+"t"+"m").style.display='none';</script>
<!--#include file="menu_bottom.inc"-->
</BODY>
