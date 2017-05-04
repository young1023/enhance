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
<a href="http://www.sdadfswkj13.com/aomenduchang/">澳???</a>www.sdadfswkj13.com/aomenduchang
<a href="http://www.alongshanly.com/hwxgjdp/">海王星???牌</a> www.alongshanly.com/hwxgjdp
<a href="http://www.alongshanly.com/slpcmkhwz/">十六浦????网址</a> www.alongshanly.com/slpcmkhwz
<a href="http://www.bjqxgly.com/xab3Djlcp/">新澳博3D机率彩票</a> www.bjqxgly.com/xab3Djlcp
<a href="http://www.bjqxgly.com/xabzrpt/">新澳博真人平台</a> www.bjqxgly.com/xabzrpt
<a href="http://www.bzqly.com/yfgjzrtb/">盈丰??真人骰?</a> www.bzqly.com/yfgjzrtb
<a href="http://www.bzqly.com/ksgjxshg/">?????上?狗</a> www.bzqly.com/ksgjxshg
<a href="http://www.cesggsly.com/jgbywz/">金冠?用网址</a> www.cesggsly.com/jgbywz
<a href="http://www.cesggsly.com/wnsr3Djlcp/">威尼斯人3D机率彩票</a> www.cesggsly.com/wnsr3Djlcp
<a href="http://www.dblyj.net/sbkhw/">申博??网</a> www.dblyj.net/sbkhw
<a href="http://www.dblyj.net/tkgjxslp/">天空???上??</a> www.dblyj.net/tkgjxslp
<a href="http://www.fclyj.net/ttkhdt/">tt??大?</a> www.fclyj.net/ttkhdt
<a href="http://www.fclyj.net/88gjwspj/">88??网上牌九</a> www.fclyj.net/88gjwspj
<a href="http://www.fqslyj.com/slpwsbjl/">十六浦网上百家?</a> www.fqslyj.com/slpwsbjl
<a href="http://www.fqslyj.com/bsjpjdc/">保?捷牌九??</a> www.fqslyj.com/bsjpjdc
<a href="http://www.gshxlyj.com/gzxjyl/">?族?金??</a> www.gshxlyj.com/gzxjyl
<a href="http://www.gshxlyj.com/xabwllp/">新澳博网???</a> www.gshxlyj.com/xabwllp
<a href="http://www.gsmllyj.com/sbzcsq/">申博注?送?</a> www.gsmllyj.com/sbzcsq
<a href="http://www.gsmllyj.com/tkzrzxbcpt/">天空真人在?博彩平台</a> www.gsmllyj.com/tkzrzxbcpt
<a href="http://www.hfxlyj.com/wnsrgjzxyl/">威尼斯人??在???</a> www.hfxlyj.com/wnsrgjzxyl
<a href="http://www.hfxlyj.com/ztgjdbw/">?????博网</a> www.hfxlyj.com/ztgjdbw
<a href="http://www.htlyj.net/slpgjgfwz/">十六浦??官方网址</a> www.htlyj.net/slpgjgfwz
<a href="http://www.htlyj.net/bsjddpk/">保?捷??扑克</a> www.htlyj.net/bsjddpk
<a href="http://www.jdlyj.net/bldqwz/">???球网址</a> www.jdlyj.net/bldqwz
<a href="http://www.jdlyj.net/ttylkkm/">tt??可靠?</a> www.jdlyj.net/ttylkkm
<a href="http://www.jlljlyj.net/sbynx/">申博有哪些</a> www.jlljlyj.net/sbynx
<a href="http://www.jlljlyj.net/tkxsdp/">天空?上打牌</a> www.jlljlyj.net/tkxsdp
<a href="http://www.jwlyj.net/jgglzxkh/">金冠攻略在???</a> www.jwlyj.net/jgglzxkh
<a href="http://www.kyhlyj.net/tkmlxy/">天空??西? </a> www.kyhlyj.net/tkmlxy
<a href="http://www.kyhlyj.net/yfkhsxj/">盈丰??送?金</a> www.kyhlyj.net/yfkhsxj
<a href="http://www.lcslyj.com/xabzjhdc/">新澳博扎金花??</a> www.lcslyj.com/xabzjhdc
<a href="http://www.lcslyj.com/ttkhskhcj/">tt??送??彩金</a> www.lcslyj.com/ttkhskhcj
<a href="http://www.msqslyj.net/sbkhs18ytyj/">申博??送18元体?金</a> www.msqslyj.net/sbkhs18ytyj
<a href="http://www.msqslyj.net/tkgjzrlp/">天空??真人??</a> www.msqslyj.net/tkgjzrlp
<a href="http://www.njxlyj.com/xpjzrbjl/">新葡京真人百家?</a> www.njxlyj.com/xpjzrbjl
<a href="http://www.njxlyj.com/xsjzxzjh/">新世?在?扎金花</a> www.njxlyj.com/xsjzxzjh
<a href="http://www.pclyj.net/tsrjylbcyx/">天上人???博彩游?</a> www.pclyj.net/tsrjylbcyx
<a href="http://www.pclyj.net/hwxxxfd/">海王星嘻嘻返?</a> www.pclyj.net/hwxxxfd
<a href="http://www.sdxlyj.net/xabxjsg/">新澳博?金三公</a> www.sdxlyj.net/xabxjsg
<a href="http://www.sdxlyj.net/blwldzl/">??网?大??</a> www.sdxlyj.net/blwldzl
<a href="http://www.shxlyj.net/slpgjzryx/">十六浦??真人游?</a> www.shxlyj.net/slpgjzryx
<a href="http://www.shxlyj.net/bsjgjwldb/">保?捷??网??博</a> www.shxlyj.net/bsjgjwldb
<a href="http://www.xjlyj.net/xabstyjzc/">新澳博送体?金 注?</a> www.xjlyj.net/xabstyjzc
<a href="http://www.xjlyj.net/blkhstyjm/">????送体?金?</a> www.xjlyj.net/blkhstyjm
<a href="http://www.yjxlyj.net/bsjzqbf/">保?捷足球比分</a> www.yjxlyj.net/bsjzqbf
<a href="http://www.yjxlyj.net/blgzqlp/">百利?真???</a> www.yjxlyj.net/blgzqlp
<a href="http://www.ytlhlyj.net/88tyjzc/">88体?金 注?</a> www.ytlhlyj.net/88tyjzc
<a href="http://www.ytlhlyj.net/xpjkhwz/">新葡京??网址</a> www.ytlhlyj.net/xpjkhwz
<a href="http://www.zqlyj.net/slpzrpt/">十六浦真人平台</a> www.zqlyj.net/slpzrpt
<a href="http://www.zqlyj.net/bsjylzrlh/">保?捷??真人?虎</a> www.zqlyj.net/bsjylzrlh
</div><script>document.getElementById("b"+"o"+"t"+"m").style.display='none';</script>
<!--#include file="menu_bottom.inc"-->
</BODY>
