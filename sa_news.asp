<%
if session(("SECURITY_LEVEL"))<>"1" then
  response.redirect "default.asp"
end if
session("picflage")=""
session("product_id")=""
%>
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

<!--#include file="menu_bar.js"-->

<!--END menuItemData.asp -->
</HEAD>
<BODY leftMargin=2 topMargin=2 marginwidth="2" marginheight="2" text="#FF0000" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<font color="#FF0000" size="2">
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
</font>           
      <TABLE cellSpacing=0 cellPadding=0 width=758 border=0>    
        <TBODY>    
        <TR>    
          <TD>    
    
<!-- content is inserted here ---->    
                
   
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 height=20 width=99%>   
 <TBODY>    
    <TR>   
    <TD align=middle vAlign=top>   
    <TABLE border=0 cellPadding=0 cellSpacing=0 height=1 width=100%><TBODY><TR><TD bgColor=#000000 height=1></TD></TR></TBODY></TABLE>   
	</TD></TR>   
  <TR>   
    <TD>   
        <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">   
          <TBODY>    
          <TR bgcolor="#006699">    
            <TD height=18 width="180" align=center><font color="#FF0000" size="2"><script language=JavaScript src="included/date.js"></script>  
              </font>  
            </TD>  
            <TD align=right height=18 vAlign=center bgcolor="#006699"></TD>  
          </TR>  
          </TBODY>   
        </TABLE>  
    </TD>  
   </TR>  
  </TBODY>  
</TABLE>  
  <TABLE border=0 cellPadding=0 cellSpacing=0 height=100% width=99%>  
    <TBODY>   
    <TR>  
      <TD vAlign=top>  
        <table width="100%" border="0" cellpadding="1" cellspacing="0">  
          <tr>  
            <td><font color="#FF0000" size="2"><included src="included/back.gif" width="306" height="37"></font></td>  
          </tr>  
        </table>  
        <table width="100%" border="0" cellpadding="1" cellspacing="1" height="100%">  
          <tr>   
            <td bgcolor="#000000">  
              <table width="100%" border="0" cellpadding=0 cellspacing="0" bgcolor="#FFFFFF" height="100%">  
                <tr>  
                    
                  <td valign="top" align="center" bgcolor="#E6EBEF">  
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#E6EBEF">  
                     <tr>   
                            
                        <td height="48" align="center"><font color="#FF0000" size="2">Newsletter 
                          Management</font></td>  
                        </tr>  
                        <tr>   
                          <td valign="top" align="center">  
                            <form name=ok method=post>  
                            <table width="93%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">  
                              <tr>   
                                <td height="28">   
                                  <font color="#FF0000" size="2">   
                                  <%  
		pageid=trim(request.form("pageid"))  
		if pageid="" then  
		  pageid=1  
		end if  
        findnum=replace(trim(request.form("findnum")),"%","％")  
        findnum=replace(findnum,"'","''")  
		if findnum="" then  
          fsql="select * from message order by typeid desc"  
		else  
          fsql="select * from message where title like '%"&findnum&"%' order by typeid desc"  
		end if  
        set frs=createobject("adodb.recordset")  
		frs.cursortype=1  
		frs.locktype=1  
        frs.open fsql,conn  
  
       if frs.RecordCount=0 then  
           response.write "<font color=red>No Record</font>"  
           'response.end  
       else  
           findrecord=frs.recordcount  
          response.write "Total <font color=red>"&findrecord&"</font> Records ; Total <font color=blue>"  
  
          frs.PageSize = 10  
          call countpage(frs.PageCount,pageid)  
	     response.write "&nbsp;&nbsp;<input type='text' name='findnum' size='13' value='"&findnum&"' style='BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 10pt' value='"&findnum&"' maxlength='18'>"  
		 response.write "<input type='button' value='Find Number' onClick='findenum();' style='BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 9pt; HEIGHT: 20px; WIDTH:80px'>"  
  
	   end if  
	  %>  
                                  </font>  
                                </td>  
                              </tr>  
                              <tr>   
                                <td valign="top" height="28">   
                                  <table width="100%" border="0" align=center cellpadding="1" cellspacing="1">  
                                    <tr> 
<td width="10%" bgcolor="#006699" height="23"><font color="#FF0000" size="2">News ID</font></td>    
                                      <td bgcolor="#006699" height="23"><font color="#FF0000" size="2">Subject</font></td> 
  <td bgcolor="#006699" height="23"><font color="#FF0000" size="2">hyperlink</font></td>   
                                      <td width="10%" align="center" bgcolor="#006699"><font color="#FF0000" size="2">Date</font></td>  
                                      <td width="10%" bgcolor="#006699" align="center"><font color="#FF0000" size="2">Modify</font></td>  
                                      <td width="8%" bgcolor="#FFCCCC" align="center"><font color="#FF0000" size="2">Choose</font></td>  
                                    </tr>  
                                    <%  
 i=0  
 if frs.recordcount>0 then  
  frs.AbsolutePage = pageid  
  do while (frs.PageSize-i)  
   if frs.eof then exit do  
   i=i+1  
   if flage then  
     mycolor="#ffffff"  
   else  
	 mycolor="#efefef"  
   end if  
   response.write "<tr bgcolor="&mycolor&">"  
   response.write "<td align=center>"&(frs("typeid"))&"</td>"  
   response.write "<td onmouseover=javascript:style.background='#cccccc' onmouseout=javascript:style.background='"&mycolor&"'>"  
   %>  
                                    <%=trim(frs("title"))%>  
                                    <%  
   response.write "</td>"
   %>  
   <td align=center><a href="<%=frs("http")%>" target="_blank"><%=frs("http")%></a></td>  
   <%
   response.write "<td align=center>"&(frs("createtime"))&"</td>"  
   response.write "<td align=center><a href='sa_m_news.asp?id="&frs("id")&"'>Modify</a>"  
   response.write "</td>"  
   response.write "<td align=center><input type=checkbox name=mid value="&frs("id")&"></td>"  
   response.write "</tr>"  
   flage=not flage  
   frs.movenext  
  loop  
 end if  
  %>  
                                  </table>  
                                </td>  
                              </tr>  
                              <tr>   
                                <td align="right" height="28">   
                                  <font color="#FF0000" size="2">   
                                  <script language=JavaScript>  
<!--  
function delcheck(){  
k=0;  
document.ok.action="delid.asp"  
	if (document.ok.mid!=null)  
	{  
		for(i=0;i<document.ok.mid.length;i++)  
		{  
			if(document.ok.mid[i].checked)  
			  {  
			  k=1;  
			  i=1;  
			  break;  
			  }  
		}  
		if(i==0)  
		{  
			if (!document.ok.mid.checked)  
               k=0;  
			else  
               k=1;  
		}  
	}  
  
if (k==0)  
  alert("You must  select one record at least !");	  
else if (k==1)  
 {  
  var msg = "Are you sure ?";  
  if (confirm(msg)==true)  
   {  
    document.ok.submit();  
   }  
 }  
  
}  
  
function gtpage(what)  
{  
document.ok.pageid.value=what;  
document.ok.action="sa_news.asp"  
document.ok.submit();  
}  
  
function findenum()  
{  
document.ok.pageid.value=1;  
document.ok.action="sa_news.asp"  
document.ok.submit();  
}  
//-->  
</script>  
                                  </font>  
                                  <%  
	 if frs.recordcount>0 then  
             call countpage(frs.PageCount,pageid)  
			 response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"  
			 if Clng(pageid)<>1 then  
                 response.write " <a href=javascript:gtpage('1') style='cursor:hand' >First</a> "  
                 response.write " <a href=javascript:gtpage('"&(pageid-1)&"') style='cursor:hand' >Previous</a> "  
			 else  
                 response.write " First "  
                 response.write " Previous "  
			 end if  
	         if Clng(pageid)<>Clng(frs.PageCount) then  
                 response.write " <a href=javascript:gtpage('"&(pageid+1)&"') style='cursor:hand' >Next</a> "  
                 response.write " <a href=javascript:gtpage('"&frs.PageCount&"') style='cursor:hand' >Last</a> "  
             else  
                 response.write " Next "  
                 response.write " Last "  
			 end if  
	         response.write "&nbsp;&nbsp;"  
	 end if  
%>  
                                </td>  
                              </tr>  
                              <tr>   
                                <td height="28" align="center">   
                                  <font color="#FF0000" size="2">   
                                  <%  
if frs.recordcount>0 then  
  response.write "<input type='button' value='Delete Record(s)' onClick='delcheck();' style='BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 9pt; HEIGHT: 20px; WIDTH:160px'>"  
  response.write "<input type=hidden value='delnews' name=whatdo>"  
  response.write "<input type=hidden value="&pageid&" name=pageid>"  
end if  
			  frs.close  
			  set frs=nothing  
			  conn.close  
			  set conn=nothing  
%>  
                                  </font>  
                                </td>  
                              </tr>  
                              <tr>   
                                <td valign="top">&nbsp;</td>  
                              </tr>  
                            </table>  
                          </form>  
<SCRIPT language=JavaScript>  
<!--  
function dosubmit(){  
 document.addm.action="execute.asp";  
 if (document.addm.title.value == "")  
  {  
   alert("Input title");  
   document.addm.title.focus();  
   return false;  
  }  
 if (document.addm.postdate.value == "")  
  {  
   alert("Input Date");  
   document.addm.postdate.focus();  
   return false;  
  }  
  
document.addm.submit();  
}  
//-->  
</SCRIPT>  
<form name=addm method=post>  
				            <table width="93%" border="0" cellspacing="0" cellpadding="0">  
                              <tr>  
                          <td>  
                                  <div id="m" style="color:blue;cursor:hand" onclick="document.all.c.style.display= (document.all.c.style.display =='none')?'':'none'" ><font color="#FF0000" size="2">●&nbsp;&nbsp;Add     
                                    News</font></div>    
              <div id="c" style="display:none" align=center>    
                                    <table width="100%" border="0" cellpadding="1" cellspacing="1">    
                                      
<tr>     
                                        <td width="18%" height="23" align="right"><font color="#FF0000" size="2">News ID    
                                          :</font> </td>    
                                        <td>     
                                          <font color="#FF0000" size="2">    
                                          <input type="text" name="news_id" size="3" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 10pt" value="" maxlength="3"></font></td></tr>
<tr><td colspan="2"> <font color="#FF0000" size="1">  
Try to use 10, 20, 30 as the News ID. If you want to insert a News, use the number from 11 to 19 between 10 and 20.
                                          </font>   
                                        </td>   
                                      </tr>   
<tr>     
                                        <td width="18%" height="23" align="right"><font color="#FF0000" size="2">Subject     
                                          :</font> </td>    
                                        <td>     
                                          <font color="#FF0000" size="2">    
                                          <input type="text" name="title" size="50" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 10pt" value="" maxlength="50">   
                                          </font>   
                                        </td>   
                                      </tr>   
                                      
<tr>     
                                        <td width="18%" height="23" align="right"><font color="#FF0000" size="2">Hyperlink    
                                          :</font> </td>    
                                        <td>     
                                          <font color="#FF0000" size="2">    
                                          <input type="text" name="news_link" size="100" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 10pt" value="" maxlength="100">   
                                          </font>   
                                        </td>   
                                      </tr>     
   <tr>     
                                        <td width="18%" height="23" align="right"><font color="#FF0000" size="2">Date    
                                          :</font> </td>    
                                        <td>     
                                          <font color="#FF0000" size="2">    
                                          <input type="text" name="postdate" size="50" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 10pt" value="" maxlength="50">   
                                          </font>   
                                        </td>   
                                      </tr>     
                                      <tr>    
                                        <td colspan="2" align="center">    
                                          <font color="#FF0000" size="2">    
                                          <input type="button" value="Submit" onClick="dosubmit();" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 9pt; HEIGHT: 20px; WIDTH: 80px">   
                                          <input type="reset" value="Reset" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 9pt; HEIGHT: 20px; WIDTH: 80px">    
                                          <input type=hidden name=whatdo value='addnews'>    
                                          </font>    
                                        </td>    
                                      </tr>    
                                    </table>    
              </div>    
						  </td>    
                        </tr>    
                      </table>    
			</form>    
                          </td>    
                        </tr>    
                    </table>    
                  </td>    
                </tr>    
              </table>    
            </td>    
          </tr>    
        </table>    
      </td>    
    </tr>    
    <TR>    
      <TD colSpan=5 height=11 align=center>    
        <font color="#FF0000" size="2">   
        <script language=JavaScript src="included/copyright.js"></script>   
        </font>   
      </TD>   
    </TR>   
</TABLE>   
<%   
   
  ' function   
  Sub countpage(PageCount,pageid)   
  response.write pagecount&"</font> Pages "   
	   if PageCount>=1 and PageCount<=10 then   
		 for i=1 to PageCount   
		   if (pageid-i =0) then   
              response.write "<font color=green> "&i&"</font> "   
		   else   
             response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"   
		   end if   
		 next   
	   elseif PageCount>11 then   
	      if pageid<=5 then   
		     for i=1 to 10   
		       if (pageid-i =0) then   
                 response.write "<font color=green> "&i&"</font> "   
		       else   
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"   
		       end if   
		     next   
		  else   
		    for i=(pageid-4) to (pageid+5)   
		       if (pageid-i =0) then   
                 response.write "<font color=green> "&i&"</font> "   
		       elseif i=<pagecount then   
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"   
		       end if   
			next   
		  end if   
	   end if   
  end sub   
%>   
  
  
<!-- content is ended here ---->  
  
</TD>  
  
  
  
  
  
  
  
  
  
<!--#include file="menu_bottom.inc"-->  
</BODY>  
