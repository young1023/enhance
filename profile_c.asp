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
    ><u>���q���</u><br>
    </font><font face="Arial" color="#000000" size="2" ><p>
      </font></strong><font  face="Arial" color="#000000" size="2"><b>���q��1982�~���ߦܤ��w��20�h�~��g��A�`���q�]�󭻴�A�u�t�]��O�W�M����j���A�@���H�~�T���D�n�~�ȡC
</p>
      <p>�ڥq�޳N�O�q�M�����O���p�A20�~�H�Ӭ�o�s�y�X�h�طs���B�@�Ϊ������ɭ�l���B���������A�`���ꤺ�~�Ȥ�߷R�C</p>
      <p>���F�ǲΪ��ݵ��~�A�̪�ڥq�P�O�o�i�h�\���l���t�C�A���Ȥᴣ�Ѧh�˪���ܡC�p��w���B�s��LCD���r���B�j�O���F�������B��l�L�����B�ӹq�{�O���B�p�g���B�o���s�i���B�p�Ƶ��F�ӳ̷s�����i�x�s�q�����USB���BFM���������B���X�m���B�������B�����������A���ӳ]�p�|���_�}�o�C�ڭ̦b�]�p�W���s���}���M�ݭn�A���_�Q�N���A�бN�դU����s���Q�k�βզX�o�l��ζǯu�ڭ̡A�ڭ̱N�ַN���դU�ĳ�</p>
      
<p>�ӧڥq�󩵦��~�Ƚd��A�{���}�o�h��LED?�~�A�p�ͦ��BLED �s�i�������A���F�ണ�Ѧ��~�~�A�٥i�����t��A�t�X�U�t�~�ȡC  
</p>
<p>�ڥq�ĥμw����i�޳N�A���~��q���V�B�����c�h�B�~���u���A�O�ۥΡB�줽�M�e§�Ϋ~�C
</p></b></font>
    </td>
  </tr>
  <tr>
    <td valign="center" width="257"><font face="Arial" color="#000000" size="2" ><strong>
    <u>�pô�Ӹ`:(Head Office)</u><br>
   �O�εo�i���q<br>
   �n�[����(����)�������q<br>
   ����E�s<br>�[��D330��<br>�¤O�f�B�j�H1��<br>
    
 �q��: 852 2798 6263 (12���u)<br>
�ǯu: 852 2751 6659 (2���u)<br>
   �q�l: <a href="mailto:info@jaguarpen.com">info@jaguarpen.com</a><br>
   ���}:  <a href="http://www.jaguarpen.com">www.jaguarpen.com</a><br>
    </font><br>

<p><font face="Arial" color="#000000" size="2" >����E�s<br>
����37-39��<br>�����Ӥ@��C�y�C<br>
    �q��: (852) 2771 1333, 2780 1029<br>
    �ǯu: (852) 2770 0351</font></p>
<p><font face="Arial" color="#000000" size="2" >�`�`��ù��Ϥ��A����1043��<br>
�p���j�H�n�y812��<br>
�q��: 0755-8225 1929<br>
�ǯu�G0755-8225 8250
</font></strong></p>

    <font face="Arial" color="#000000" size="2" ><strong><u>�D�n�p���H:</u><br>
    </strong></font>
      <table border="0" cellpadding="3" cellspacing="0" width="100%">
        <tr>
          <td width="80%"  ><font  face="Arial" color="#000000" size="2"><b>Mr.
            Andy Ma<br>���Ӱ����<br>
�q�l: <a href="mailto:andy@jaguarpen.com">andy@jaguarpen.com</a><br>
     ������852 91936592  <br>������ 13068461233
</b></font><br></td>
          <td width="20%"  valign="top" ><font  face="Arial" color="#000000" size="2"><b>����</b></font></td>
        </tr>
        <tr>
          <td width="80%"  ><font  face="Arial" color="#000000" size="2"><b>Mr.
            Andrew Chang<br>�i�l�����<br>
     �q�l: <a href="mailto:andrew@jaguarpen.com">andrew@jaguarpen.com</a> <br>
     ������852 91983489 <br>������ 13822883284
</b></font><br></td>
          <td width="20%" valign="top"  ><font  face="Arial" color="#000000" size="2"><b>����</b></font></td>
        </tr>
        <tr>
          <td width="80%"  ><font  face="Arial" color="#000000" size="2"><b>Mr.
            Jimmy Tsang <br>���w������<br>�q�l:<a href="mailto:jimmy@jaguarpen.com"> jimmy@jaguarpen.com</a><br>
     ������852 91982542 <br> ������ 13822830023
</b></font></td>
          <td width="20%" valign="top"  ><font  face="Arial" color="#000000" size="2"><b>����</b></font><br></td>
        </tr>
        <tr>
          <td  height="21"><font  face="Arial" color="#000000" size="2"><b>Mr.
            Tsang Tak Shan<br>���w������<br>�q�l:<a href="mailto:info@jaguarpen.com"> info@jaguarpen.com</a><br>������<br>852 94690703</b></font><br></td>
          <td valign="top" height="21"><font  face="Arial" color="#000000" size="2"><b>����</b></font></td>
        </tr>
      </table>
      <p>
    <font face="Arial" color="#000000" size="2" ><strong><br>
    </strong></font><br>
    <font face="Arial" color="#000000" size="2" ><strong><u>�~������</u><br>
    �s�y��<br>
    �X�f��<br>
    �W�a�N�z��<br>
    <br>
    <u>OEM �A�ȨѰ�</u><br>
    �O<br>
    <u><br>
    �u�{�v�ƶq<br>
    </u>2<br>
    <br>
    <u>���q�п�~��</u><br>
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
    <u>�X�f����</u><br>
    ���@��<br>
    <br>
    <u>���u�ƶq</u><br>
    134<br>
    <br>
    <u>�C��Ͳ��e�q</u><br>
    800,000pcs<br>
    <br>
    <u>Product Range</u><br>
    Ball pen, roller pen, fountain pen, pencil &amp; gift boxes<br>
    <br>
    <u>Factory size in meter</u><br>
    6000 square meter<br>
    <br>
    <u>�u�t:</u><br>
    <u>����u�t</u><br>
    �n�[����(����)�������q�s�F��<br>
 �鶧���s��#324����<br>
 �M�����q�v�d�Ǯհ�
    
    
    <br>
    <u>�O�W�u�t :</u><br>
   Jaguar Pen (Taiwan) Co Ltd<br>
�O�_�T����<br>���R��315��24��
    
    <br>
    <u>���u�Ա�</u><br>
    �Ͳ������u&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 100<br>
    QC ���u
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    10<br>
    R &amp; D ���u
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    4<br>
    �u�{�v 
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;
    2<br>
    �����ƳB &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 18</strong></font></p>
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