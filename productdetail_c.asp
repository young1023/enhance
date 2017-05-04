<!--#include file="database.inc"-->
<%
strpid=request("id")
session("prid")=strpid
'response.write strpid
'response.write session("prid")
isfind=false
if trim(strpid)<>"" then
	strsqldet="SELECT product.cdescription,product.ID, product.cModel, category.cCategory,  " _
			 &"product.moneytype,product.Pic2Path, " _
			 &"product.cpacking, product.minorder, product.cdelivery, exportprice.exportprice, " _
			 &"product.cterm, product.sampleprice " _
			 &"FROM (product LEFT JOIN category ON product.category = category.Id) " _
			 &"LEFT JOIN exportprice ON product.exportprice = exportprice.Id " _
			 &"Where product.id="& strpid

	set acres=conn.execute(strsqldet)
	if not acres.eof then
		isfind=true
		strmodel=trim(acres("cmodel")&"")
		strcategory=trim(acres("ccategory")&"")
		strdes=trim(acres("cdescription"))
		strpic=acres("pic2Path")
		strpak=trim(acres("cpacking")&"")
		strorder=trim(acres("minorder")&"")
		strdelivery=trim(acres("cdelivery")&"")		
		strexprice=trim(acres("exportprice")&"")				
		strterm=trim(acres("cterm")&"")						
		strsampleprice=trim(acres("sampleprice")&"")
		strmoneytype=trim(acres("moneytype")&"")
	
	
	end if 
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="description"
content="manufacturer of pens, writing instruments, stationery, gifts, ball pens, premiums &amp; incentives, jaguar pens giftware &amp; novelties,  pens &amp; pen sets">
<meta name="keywords"
content="manufacturer, pens, writing instruments, WDC, stationery, gifts, ball pens, premiums &amp; incentives, jaguar pens giftware &amp; novelties, Winson, pens &amp; pen sets">

<TITLE>Jaguar Pen Limited</TITLE>
<LINK title=style1 
href="images/basic.css" type=text/css rel=stylesheet><!--BEGIN menuItemData.asp -->
<SCRIPT language=JavaScript 
src="images/js_dynMenuFunctions.js">
<!-- Functions used to create the dynamic HTML drop-down menus -->
</SCRIPT>

<!--#include file="menu_bar_c.js"-->

<!--END menuItemData.asp -->
<script language="JavaScript">
<!--
function doorder()
{
	var strmsg;
	var str1;
	strmsg="";	
	if(document.main.txtt.value=="")
	strmsg=strmsg+"Please input a number.\n";
		if(isNaN(document.main.txtt.value))
			strmsg=strmsg+"Please input a number.\n";
		if((document.main.txtt.value)<=0)
		     strmsg="Please input again.\n";
		if (document.main.txtt.value.indexOf(".")!=-1)				 
		strmsg="The number must be an integer.\n";
				
	if(strmsg!="")
		alert(strmsg);
	else
	{
	 
		document.main.submit();
	}

}

//-->
</script>
<SCRIPT language=javascript>
function dosubmit()
{
 var i,j,k,strmsg;
 k=0;
 j=0;
 strmsg="";
 for(i=0;i<document.form1.txtt.length;i++)
  { 
   j=1;
   if(document.form1.txtt(i).value=="")
     k++;
   else 
    {
     if(isNaN(document.form1.txtt(i).value))
       {strmsg="Please input number.\n";
         break;
       }
     if((document.form1.txtt(i).value)<=0)
       {strmsg="Please input again.\n";
		break;
       }
     if (document.form1.txtt(i).value.indexOf(".")!=-1)
       {strmsg="The number must be an integer!.\n";
        break;
       }
     }
  }

if(i>0 && k==document.form1.txtt.length)
  strmsg="Please input number.\n";
  //if(i>0&& j==document.form1.txtt.length)
  //strmsg="Please enter a quantity of one or more!.\n";
if(i==0 && j==0)
{
 if (document.form1.txtt.value=="")			    
  strmsg="Please input number.\n";
 else
  { 
   if(isNaN(document.form1.txtt.value))
     strmsg="Please input number.\n";					  
   if((document.form1.txtt.value)<=0)					 
	 strmsg="Please input again!.\n";			   
   if (document.form1.txtt.value.indexOf(".")!=-1)			 
     strmsg="The number must be an integer !\n";
  }		  
}

if (strmsg!="")
  alert(strmsg);
else
{
 document.form1.submit();
}

}
//-->
</SCRIPT>
</HEAD>
<BODY bgColor=#ffffff leftMargin=2 topMargin=2 marginwidth="2" marginheight="2" text="#FF0000" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<font face="Arial" >

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
            <p align="center">
 
 
 
<table border="0" cellspacing="0" cellpadding="0" width="100%"> 
  
   <tr> 
   
    <td width="642" valign="top">  
      <div align="center"> 
        <%if isfind then%> 
	   <form name="main" action="order_c.asp" method="post"> 
          <font color="#FF0000"  face="Arial">  
                  <input type="hidden" name="txtpoid" value="<%= strpid %>"> 
                  </font> 
          <table border="0" cellspacing="5" cellpadding="5" bgcolor="#FFFFFF" width="640"> 
            <tr>  
              <td width="363" bgcolor="#FFFFFF" rowspan="8">  
                               <div align="center"><font color="#FF0000"  face="Arial">  
                    
                  <%if not isnull(strpic) then%> 
				  <img src="showimg.asp?id=<%=strpid%>&flage=2" width="480" height="360" border="0"> 
                  <% Else  %>   
                  No Photo    
                  <%end if%> 
                   
                  </font></div> 
  
 
              </td> 
              <td bgcolor="#FFFFFF" height="19" align="left" width="59">  
                <div align="left"><font color="#FF0000"  face="Arial">型號:</font></div> 
            </td> 
              <td bgcolor="#FFFFFF" align="left" width="198">   <font color="#FF0000"  face="Arial"><%=strmodel%></font></td> 
            </tr> 
            <tr bgcolor="#003366">  
              <td bgcolor="#FFFFFF" align="left" width="59">  
                <div align="left"><font color="#000000" face="Arial" >類型:</font></div> 
             </td> 
              <td bgcolor="#FFFFFF" align="left" width="198">  <font color="#000000" face="Arial" ><%=strcategory%></font></td> 
            </tr> 
            <tr bgcolor="#003366">  
              <td bgcolor="#FFFFFF" align="left" width="59">  
                <div align="left"><font color="#000000"  face="Arial">內容:</font></div> 
              </td> 
              <td bgcolor="#FFFFFF" align="left" width="198">  
               <div align="left"> <font color="#000000"  face="Arial"><%=strdes%></font></div> 
 
              </td> 
            </tr> 
             <tr bgcolor="#FFFFCC">  
              <td bgcolor="#FFFFFF" align="left" width="59">  
                <div align="left"><font color="#000000"  face="Arial">包裝:</font></div> 
              </td> 
              <td bgcolor="#FFFFFF" align="left" width="198">  
               <font color="#000000"  face="Arial"><%=strpak%></font> 
              </td> 
            </tr> 
            <tr bgcolor="#FFFFCC">  
              <td bgcolor="#FFFFFF" height="22" align="left" width="59">  
                <div align="left"> 
                  <p align="left"><font color="#000000"  face="Arial">最少定購:</font></div>        
              </td> 
              <td bgcolor="#FFFFFF" height="22" align="left" width="198">  
                <font color="#000000"  face="Arial"><%=strorder%></font>       
              </td> 
            </tr> 
            <tr bgcolor="#FFFFCC">  
              <td bgcolor="#FFFFFF" align="left" width="59">  
                <div align="left"><font color="#000000"  face="Arial">交貨:</font></div> 
              </td> 
              <td bgcolor="#FFFFFF" align="left" width="198">  
               <font color="#000000"  face="Arial"><%=strdelivery%></font> 
              </td> 
            </tr> 
            <tr bgcolor="#FFFFCC">  
              <td bgcolor="#FFFFFF" align="left" width="59">  
          <div align="left"><font color="#000000"  face="Arial">價格:</font></div> 
           
              </td> 
              <td bgcolor="#FFFFFF" align="left" width="198"> <font  face="Arial">  <% if strexprice="" then%>
                <font color="#000000"> Contact Us</font>    
			<%else%>     
                    
                </font> 
				     <div align="left"><font color="#000000"  face="Arial"><%=strexprice%></font></div> 
		    <%end if%> 
               </td> 
            </tr> 
            <tr bgcolor="#FFFFCC">  
              <td bgcolor="#FFFFFF" align="left" width="59">  
                <font color="#000000"  face="Arial">條款:</font> 
              </td> 
              <td bgcolor="#FFFFFF" align="left" width="198">  
                <font color="#000000"  face="Arial"><%=strterm%></font> 
              </td> 
            </tr> 
          
            <tr bgcolor="#FFFFCC">  
              <td width="363" bgcolor="#FFFFFF">  
                <div align="left"><font color="#000000"  face="Arial">樣本價錢:  
                  <%if trim((strsampleprice)&"") =""  then%> 
                  聯絡我們 
                  <%else%> 
                  <%=strmoneytype&strsampleprice%>   
                  <%end if%>  
                   
                  </font> </div> 
              </td> 
              <td width="263" bgcolor="#FFFFFF" align="center" colspan="2">  
                <%if trim((strsampleprice)&"") =""  then%> 
                 &nbsp;
                  <%else%> 
                 <div align="left"><font color="#000000"  face="Arial">要求報價</font></div>   
                 <%end if%>
               
              </td> 
            </tr> 
         
            <tr bgcolor="#FFFFCC">  
               
              <td width="363" bgcolor="#FFFFFF" align="center">   <%if trim((strsampleprice)&"") =""  then%> 
                  &nbsp;
                  <%else%> 
                
 
                <div align="left"><font color="#000000"  face="Arial">數量 :&nbsp;&nbsp;&nbsp;&nbsp; 
                  <input type="text" name="txtt" maxlength="12"> 
                    <input type="hidden" name="pcid" value="">   
                   
                  </font> 
                 <br><br>
                   <div align="left"><a href="javascript:doorder();"><font color="#FF0000"  face="Arial">加入樣板籃</font></a></div>  
              <input type="hidden" name="pcid" value="">    
     <%end if%>  
                </div>
               
              </td> 
</form>
	   <form name="form1" action="enquiry_c.asp" method="post"> 
   

              <td width="263" bgcolor="#FFFFFF" colspan="2">  
                  
                  <font face="Arial" font color="#000000">數量:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                  <input type="text" name="txtt" maxlength="12"> 
                   
 <input type="hidden" name="txtpoid" value="<%= trim(Acres("id")&"") %>"><br>

                  <br> 
                   
                  </font> 
                <div align="left"><a href="javascript:dosubmit();"><font color="#FF0000"  face="Arial">加入要求報價表</font></a></div>  
              </td> 
            </tr> 
<tr>
   <td colspan="3" align="center">             
<A HREF="mailto:?subject=Jaguarpen 的筆產品&body=請看看  http://www.jaguarpen.com/productdetail_c.asp?id=<%= trim(Acres("id")&"") %>">
電郵這頁資料給同事</A>		 
   </td>             
</tr>             
		 
             
			 
          </table> 
	 
</form> 
		<%end if%>   
                  
      </div> 
    </td> 
  </tr> 
</table> 
 
 
</TD> 
        <!--#include file="menu_bottom_c.inc"--> 
</body>   
