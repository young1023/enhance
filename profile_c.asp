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

<!--#include file="menu_bar_c.js"-->

<!--END menuItemData.asp -->
</HEAD>
<BODY bgColor=#ffffff leftMargin=2 topMargin=2 marginwidth="2" marginheight="2" text="#000000" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
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
            
  <center>
 <table border="0" width="637">
  <tr>
    <td valign="top" colspan="2" width="596"><strong><font face="Arial" color="#000000" color="#000000" size="2"
    ><u>公司資料</u><br>
    </font><font face="Arial" color="#000000" size="2" ><p>
      </font></strong><font  face="Arial" color="#000000" size="2"><b>本司自1982年成立至今已有20多年制筆經驗，總公司設於香港，工廠設於臺灣和中國大陸，一直以外貿為主要業務。
</p>
      <p>我司技術力量和資金實力雄厚，20年以來研發製造出多種新型、耐用的中高檔原子筆、鋼筆等等，深受國內外客戶喜愛。</p>
      <p>除了傳統金屬筆外，最近我司致力發展多功能原子筆系列，給客戶提供多樣的選擇。如行針表筆、新款LCD跳字表筆、強力馬達按摩筆、原子印章筆、來電閃燈筆、雷射筆、發光廣告筆、計數筆；而最新的有可儲存電腦資料USB筆、FM收音機筆、六合彩筆、錄音筆、香水筆等等，未來設計會不斷開發。我們在設計上的連續突破仍然需要你的寶貴意見，請將閣下任何新的想法或組合發郵件或傳真我們，我們將樂意為閣下效勞</p>
      
<p>而我司更延伸業務範圍，現正開發多樣LED?品，如匙扣、LED 廣告筆等等，除了能提供成品外，還可供應配件，配合各廠業務。  
</p>
<p>我司採用德國先進技術，產品質量卓越、種類繁多、外形優雅，是自用、辦公和送禮佳品。
</p></b></font>
    </td>
  </tr>
  <tr>
    <td valign="center" width="257"><font face="Arial" color="#000000" size="2" ><strong>
    <u>聯繫細節:(Head Office)</u><br>
   力佳發展公司<br>
   積架金筆(香港)有限公司<br>
   香港九龍<br>觀塘道330號<br>威力貨運大廈1樓<br>
    
 電話: 852 2798 6263 (12條線)<br>
傳真: 852 2751 6659 (2條線)<br>
   電郵: <a href="mailto:info@jaguarpen.com">info@jaguarpen.com</a><br>
   網址:  <a href="http://www.jaguarpen.com">www.jaguarpen.com</a><br>
    </font><br>

<p><font face="Arial" color="#000000" size="2" >香港九龍<br>
花園街37-39號<br>金旺樓一樓C座。<br>
    電話: (852) 2771 1333, 2780 1029<br>
    傳真: (852) 2770 0351</font></p>
<p><font face="Arial" color="#000000" size="2" >深圳市羅湖區文錦中路1043號<br>
聯興大廈南座812室<br>
電話: 0755-8225 1929<br>
傳真：0755-8225 8250
</font></strong></p>

    <font face="Arial" color="#000000" size="2" ><strong><u>主要聯絡人:</u><br>
    </strong></font>
      <table border="0" cellpadding="3" cellspacing="0" width="100%">
        <tr>
          <td width="80%"  ><font  face="Arial" color="#000000" size="2"><b>Mr.
            Andy Ma<br>馬志堅先生<br>
電郵: <a href="mailto:andy@jaguarpen.com">andy@jaguarpen.com</a><br>
     香港手機852 91936592  <br>中國手機 13068461233
</b></font><br></td>
          <td width="20%"  valign="top" ><font  face="Arial" color="#000000" size="2"><b>董事</b></font></td>
        </tr>
        <tr>
          <td width="80%"  ><font  face="Arial" color="#000000" size="2"><b>Mr.
            Andrew Chang<br>張子科先生<br>
     電郵: <a href="mailto:andrew@jaguarpen.com">andrew@jaguarpen.com</a> <br>
     香港手機852 91983489 <br>中國手機 13822883284
</b></font><br></td>
          <td width="20%" valign="top"  ><font  face="Arial" color="#000000" size="2"><b>董事</b></font></td>
        </tr>
        <tr>
          <td width="80%"  ><font  face="Arial" color="#000000" size="2"><b>Mr.
            Jimmy Tsang <br>曾德明先生<br>電郵:<a href="mailto:jimmy@jaguarpen.com"> jimmy@jaguarpen.com</a><br>
     香港手機852 91982542 <br> 中國手機 13822830023
</b></font></td>
          <td width="20%" valign="top"  ><font  face="Arial" color="#000000" size="2"><b>董事</b></font><br></td>
        </tr>
        <tr>
          <td  height="21"><font  face="Arial" color="#000000" size="2"><b>Mr.
            Tsang Tak Shan<br>曾德珊先生<br>電郵:<a href="mailto:info@jaguarpen.com"> info@jaguarpen.com</a><br>香港手機<br>852 94690703</b></font><br></td>
          <td valign="top" height="21"><font  face="Arial" color="#000000" size="2"><b>董事</b></font></td>
        </tr>
      </table>
      <p>
    <font face="Arial" color="#000000" size="2" ><strong><br>
    </strong></font><br>
    <font face="Arial" color="#000000" size="2" ><strong><u>業務類型</u><br>
    製造商<br>
    出口商<br>
    獨家代理商<br>
    <br>
    <u>OEM 服務供商</u><br>
    是<br>
    <u><br>
    工程師數量<br>
    </u>2<br>
    <br>
    <u>公司創辦年份</u><br>
    1982<br>
    <br>
    <u>Type of Plant &amp; Machinery in Factories</u><br>
    Automatic assembling machines<br>
    Automatic lathe machines<br>
    Computerized cutting machines<br>
    Electric spark forming machine<br>
    Electroplating equipment<br>
    Heat transfer printing machines<br>
    Hot stamping machines<br>
    Lacquering / Painting system from USA and Japan<br>
    Plastic injection machines<br>
    Precision surface grinders<br>
    Screen-printing machines<br>
    Vibratory finishing machines<br>
    <br>
    <u>出口市場</u><br>
    全世界<br>
    <br>
    <u>員工數量</u><br>
    134<br>
    <br>
    <u>每月生產容量</u><br>
    800,000pcs<br>
    <br>
    <u>Product Range</u><br>
    Ball pen, roller pen, fountain pen, pencil &amp; gift boxes<br>
    <br>
    <u>Factory size in meter</u><br>
    6000 square meter<br>
    <br>
    <u>工廠:</u><br>
    <u>中國工廠</u><br>
    積架金筆(中國)有限公司廣東省<br>
 潮陽市廣汕#324公路<br>
 和平路段師範學校側
    
    
    <br>
    <u>臺灣工廠 :</u><br>
   Jaguar Pen (Taiwan) Co Ltd<br>
臺北三重市<br>仁愛街315巷24號
    
    <br>
    <u>員工詳情</u><br>
    生產部員工&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 100<br>
    QC 員工
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    10<br>
    R &amp; D 員工
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    4<br>
    工程師 
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;
    2<br>
    香港辦事處 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 18</strong></font></p>
    </td>
    <td valign="top" width="350"><strong><font face="Arial" color="#000000" size="2" >
    <img
    src="images/profile/profile01.jpg" alt="profile01.jpg (20903 bytes)" width="300" height="193"><br>
    <br>
    <img src="images/profile/profile02.jpg" alt="profile02.jpg (23584 bytes)" width="300" height="193"></font></strong><br>
    <br>
    <strong><font face="Arial" color="#000000" size="2" >
    <img
    src="images/profile/profile03.jpg" alt="profile03.jpg (24045 bytes)" width="300" height="193"></font></strong></td>
  </tr>
  
</table>
</center></div>




<!---- Main Content is ended here ------>  

</TD>
<!--#include file="menu_bottom_c.inc"-->
</BODY>