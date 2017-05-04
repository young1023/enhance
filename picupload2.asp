<!--#include file="database.inc"-->
<% 
if session("SECURITY_LEVEL")="" then
	response.redirect "login.asp"
end if
strprodid=session("product_id")
if strprodid<>"" then
	session("product_id")=strprodid
	session("picflage")=2
	set rec=server.createobject("ADODB.recordset")
    rec.Open "SELECT pic2path FROM product where id=" & strprodid & " ",conn,1,3
    if not rec.eof then
     pic=rec("pic2path")
	 'response.write "<font color=red>ok</font>"
    end if
    rec.close
    set rec=nothing
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="imagetoolbar" content="no">
<TITLE>Enhance Technologies</TITLE>
<LINK title=style1 
href="images/basic.css" type=text/css rel=stylesheet><!--BEGIN menuItemData.asp -->
<SCRIPT language=JavaScript 
src="images/js_dynMenuFunctions.js">
<!-- Functions used to create the dynamic HTML drop-down menus -->
</SCRIPT>
<script language="JavaScript">
<!--
function chkfile()
{
	if((document.picupload.picfile1.value==""))
	alert("Please select a picture file!");
	else
	document.picupload.submit();
}
//-->
</script>

<!--#include file="menu_bar.inc"-->

<!--END menuItemData.asp -->
</HEAD>
<BODY bgColor=#ffffff leftMargin=2 topMargin=2 marginwidth="2" marginheight="2" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<font color="#FF0000">
<p>
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
      border=0 name=DCSIMG> </p>
</NOSCRIPT>         
</font>         
      <TABLE cellSpacing=0 cellPadding=0 width=758 border=0>  
        <TBODY>  
        <TR>  
          <TD>  
  
<!-- content is inserted here ---->  
<table border="0" cellspacing="0" cellpadding="0" width="780" height="360">  
  <tr>  
    <td width="644" valign="middle">   
      <div align="center">  
        <p><a href="DatAdmin.asp"><font face="Arial, Helvetica, sans-serif" size="2">Administrator Management</font></a></font></p>  
        <p>Upload the picture here:<br></p> 
        <font color="#FF0000"> 
	  <%if not isnull(pic) then %> 
          <img src="showimg.asp?id=<%=strprodid%>&flage=2" border=1 width="163" title="small picture"> 
	  <%else  
	      response.write "No Big Photo !"  
		end if  
	%>  
        </font>  
        <form name="picupload" enctype="multipart/form-data" method="post" action="uploadresult.asp">  
          <table width="80%" border="0" cellspacing="1" cellpadding="4" bgcolor="#999999">  
            <tr>   
              <td height="23" bgcolor="#FFFFFF">Stage 2</td>  
              <td height="23" bgcolor="#FFFFFF">  
                <a href="picupload.asp?picflage=1">Upload   
                  a small picture</a></div>  
              </td>  
            </tr>  
            <tr>  
              <td bgcolor="#FFFFFF">Add a big picture</td>  
              <td bgcolor="#FFFFFF"><b>  
                <font color="#FF0000"> 
                <input type="file" name="picfile1"> 
                </font> 
                </b></td> 
            </tr> 
          </table> 
          <b><font color="#FF0000"><br> 
          <input type="button" name="Submit2" value="Upload Picture" onClick="chkfile();"> 
          <input type="reset" name="Reset" value=" Reset Form">  
           <input type="hidden" name="product_id" value="<%=strprodid%>">  
		  <%if not isnull(pic) then  
          response.write "<font color='#CC0000' size='2'><a href='delpic.asp'><font color='#FF0000'>"  
		  response.write "Del The Big Picture</font></a></font>"  
		  end if  
		  %>  
          </font>  
		  </b>   
        </form>  
            <a href="editprod.asp?product_id=<%=strprodid%>">Continueous to edit product</a></div> 
    </td> 
  </tr> 
</table> 

<!-- content is ended here ----> 
 
</TD> 

 
<!--#include file="menu_bottom.inc"--> 
</BODY> 
