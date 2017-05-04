<!--#include file="database.inc"-->
<% 

response.expires=0 
straction=request("action_button")
if straction<>""then
  session("where")=""
end if

'response.write straction
  if straction=""then
    straction="model"
  end if
stractt=request("action")
 
pageno=trim(request("pageno"))
strwhere=trim(request("strwhere"))
if pageno="" then
	pageno=1
	strwhere=""
end if
strsearchsql=""


	strsearchsql="SELECT product.description, product.new, product.Model, exportprice.Exportprice, category.Category, " _
	      &"product.pic2path, product.id " _
		  &"FROM category RIGHT JOIN (exportprice RIGHT JOIN " _
		  &"product ON exportprice.Id = product.exportprice) " _
		  &" ON category.Id = product.category " _
		  &"where product.new =1 "

  'response.write strsearchsql


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="imagetoolbar" content="no">
<TITLE>Jaguar Pen Limited</TITLE>
<LINK title=style1 
href="images/basic.css" type=text/css rel=stylesheet><!--BEGIN menuItemData.asp -->
<SCRIPT language=JavaScript 
src="images/js_dynMenuFunctions.js">
<!-- Functions used to create the dynamic HTML drop-down menus -->
</SCRIPT>


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
       {strmsg="請輸入數量.\n";
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
  strmsg="請輸入數量.\n";
  //if(i>0&& j==document.form1.txtt.length)
  //strmsg="Please enter a quantity of one or more!.\n";
if(i==0 && j==0)
{
 if (document.form1.txtt.value=="")			    
  strmsg="請輸入數量.\n";
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


<!--#include file="menu_bar_c.js"-->


<!--END menuItemData.asp -->
</HEAD>
<BODY bgColor=#ffffff leftMargin=2 topMargin=2 marginwidth="2" marginheight="2" text="#FF0000" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
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



<!-- START OF Data Collection Server TAG --><!-- Copyright 2002 NetIQ Corporation --><!-- V2.1 -->
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



                                           
      <TABLE cellSpacing=0 cellPadding=0 width=780 border=0>
        <TBODY>
        <TR>
          <TD>
            <p align="center">
  <table border="0" cellspacing="0" cellpadding="0" width="780">
    <tr> 
      <td valign="top" width="790" align="center"> 
   			<% If strsearchsql<>"" then %>
            
            <b>  <font size="1">                              
            <% 
					i=1
					j=0
					adopenkeyset=3
					adlockoptimistic=2
					set acres=nothing
					Set acres = createobject("ADODB.Recordset")
					acres.CursorType=adOpenKeyset
					acres.LockType=adLockOptimistic
					acres.open strsearchsql,conn
					recount=acres.recordcount
				'response.write recount
					if pageno>1 then
						i=(pageno-1)*9
						acres.move i
					end if
					if not acres.eof then
				%></font>
</b> 
				<% if recount=1 then%>           
				  <br><font size="1"> 
           找到 <%=recount%> 個最新產品.</font> 
            <%else%>		                   
				<br><font size="1"> 找到 <%=recount%> 個最新產品.</font> 
                  
		    <%end if %>      
		   <form name="form1" action="enquiry_c.asp" method="post">
            <table width="98%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
              <tr  > 
                <td> 
                  <div align="left"> <font size="2"> 
                    <%call showpageno(pageno,straction,strwhere) 
					  oldstrwhere=strwhere
					 %>
                    </font> </div>
                </td>
              </tr>
              <tr> 
                  <td width="17%"> 
                    <div align="left"><font size="2"> 
                      <% If instr(1,strwhere,"order by")>0 then
						strwhere=mid(strwhere,1,instr(1,strwhere,"order")-1)
					end if
				  		strwhere=strwhere & " order by product.Model"
				  %>
                       <a href="snew_c.asp?action_button=<%= server.URLencode(straction) %>&strwhere=<%= server.URLencode(strwhere) %>&pageno=<%= 1 %>"> 
                      <img src="images/product/up.gif" width="15" height="15" border="0"></a>  
                      型號 
                      <% If instr(1,strwhere,"order by")>0 then
						strwhere=mid(strwhere,1,instr(1,strwhere,"order")-1)
					end if
				  		strwhere=strwhere & " order by product.Model DESC "
				  %>
                      <a href="snew_c.asp?action_button=<%= server.URLencode(straction) %>&strwhere=<%= server.URLencode(strwhere) %>&pageno=<%= 1 %>">   
                      <img src="images/product/down.gif" width="15" height="15" border="0"></a> 
                      </font></div>
                </td>
                
                
              </tr>
</table>
<table width="98%" border="0" cellspacing="5" cellpadding="2" bgcolor="#FFFFFF">
         <tr > 
              <% 	  	do while (not acres.eof)
						if j>8 then exit do
				  %>
             
              <td width="15%" align="center" bgcolor="#FFFFFF" bordercolor="red">
 <a href="productdetail_c.asp?id=<%=trim(acres("id")&"") %>"> 
                  <font size="2"> 
                  <%if not isnull(acres("pic2path")) then%>
                  <img src="showimg.asp?id=<%=acres("id")%>&flage=2" border=1 width="125" height="100">
                  <%else  
                      response.write "<font size=2>No Photo</font>"
                    end if%>
                  </font>
                    </a> <br>
                <a href="productdetail_c.asp?id=<%=trim(acres("id")&"")%>">
                  <font size="-2">
				  <%=acres("model")%></font></a><br>

<font size="-2">感興趣數量</font>
<input type="text" name="txtt" size="5" maxlength="5">
 <input type="hidden" name="txtpoid" value="<%= trim(Acres("id")&"") %>">
</td>
                 <td width="15%" valign="top" align="left" bgcolor="#FFFFFF"><font size="-2" color=#000000><%=trim(acres("description")&"") %></font><br>


       </td>
                 
                
              
              <%  		acres.movenext
						j=j+1 
  If j = 3 or j = 6 then %> 
</tr> 
<% end if 
				    loop
				  %>
</tr>
              <tr > 
                <td colspan="10"> 
                      <div align="center"> 
                        <font size="2"> 
                        <input type="button" value="加入要求報價表" onClick="javascript:dosubmit();" name="button">
                        <%  
						 if trim(session("enquiryno"))<>"" then %>
                        <input type="button" name="Button" value="查諮詢籃" onClick="javascript:location='enquiry_c.asp'">  
						  <%  
				          strsql="Select * from Acclist where A_id= "& session("enquiryno") &"  "
				           set acresorder=conn.execute(strsql)				
				          if not Acresorder.eof then %>
				         
                        <input type="button" name="Button2" value="提交" onClick="javascript:location='sendto.asp'">  
                        <input type="button" name="Button3" value="清空諮詢籃" onClick="javascript:location='deletecar.asp'">  
                        <%end if  
					       set acresorder=nothing
					      set strsql=nothing
					    %>
                        <input type="hidden" name="hidbuttion" value="">  
                        <%end if%>  
                        </font>
                      </div>
                      <div align="right"> 
                      <font size="2"> 
                      <%call showpageno(pageno,straction,oldstrwhere) %>
                      </font> </div>
                </td>
              </tr>
            </table> 
            </form>				
            <% Else  %>
<table width="98%" border="0" cellspacing="5" height=200 cellpadding="2" bgcolor="#FFFFFF">
         <tr > 
              <td align="center">
              
            <font size="2">沒有紀錄</font>.
              </td>
         <tr>
</table>
            <% End If %>                                     
            <% End If %>                                       
           
      </td>
    </tr> 
  </table> 
 </div>  
</TD>
 <!--#include file="menu_bottom_c.inc"-->

</body>  
<% 
function showpageno(pageno,straction,strwhere) 
	if recount>9 then 
		lastpage=recount\9 
		if recount mod 9 >0 then 
			lastpage=lastpage+1 
		end if 
		if pageno>9 then 
		     response.write "<a href='snew_c.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&pageno=1'> The First Page</a>&nbsp;&nbsp;" 
			response.write "<a href='snew_c.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&pageno="&(pageno-9-(pageno  mod 9) )&"'>Previous </a>&nbsp;&nbsp;" 
		end if 
		strtemp=pageno 
		if (pageno Mod 9 )=0 then 
		   strtemp=strtemp-9 
		end if 
	 for i=(strtemp-(strtemp mod 9)+1) to (strtemp+9-(strtemp mod 9)) 
	         if lastpage<i then  exit for  
            if i- pageno=0 then 
				response.write cstr(i)&"&nbsp;&nbsp;" 
			else 
				response.write "<a href='snew_c.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;" 
			end if	 
		next 
		if (pageno\9)<(lastpage\9) then 
		       	response.write "<a href='snew_c.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&Pageno="&(pageno+11-(pageno mod 9)) &"'>Next</a>&nbsp;&nbsp;" 
			   response.write "<a href='snew_c.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&Pageno="& (lastpage) &"'> The Last Page</a>&nbsp;&nbsp;" 
		end if 
		 
 end if 
end function 
%> 












