<!--#include file="database.inc"-->
<%
'===================
'
' select offer
'
'===================

  id=split(trim(request.form("product_id")),",")
 
  zid=split(trim(request.form("zero")),",")

 for i=0 to Ubound(zid)
        sql="Update product set new='0' where id="&zid(i)
  conn.execute(sql)
 next

'if id <> "" then
 for i=0 to Ubound(id)
        sql="Update product set new='1' where id="&id(i)
  conn.execute(sql)
 next

'end if

pageno=trim(request("pageno"))
strwhere=trim(request("strwhere"))
if session(("SECURITY_LEVEL"))<>"1" then
  response.redirect "default.asp"
end if
if pageno="" then
	pageno=1
	strwhere=""
end if%>
<!--#include file="database.inc"-->
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
<script language="JavaScript">
<!--
function doedit()
{
	var i,j,k,strmsg;
	k=0;
	strmsg="";
	if (document.editproduct.product_id!=null)
	{
		for(i=0;i<document.editproduct.product_id.length;i++)
		{
			if(document.editproduct.product_id[i].checked)
			k++;
		}
		if(i>0 && k==0)
		strmsg="You have to select one product";
		if (k>1)
		strmsg="You can only select one product"
		if(i==0)
		{
			if (!document.editproduct.product_id.checked)
			strmsg="You have to select one product to edit";
		}
	}
	if (strmsg!="")
	alert(strmsg);
	else
	{
		document.editproduct.submit();
	}

}

function seachmodel()
{
document.editproduct.action="product_list.asp";
document.editproduct.submit();
}

//-->
</script>

<!--#include file="menu_bar.inc"-->

<!--END menuItemData.asp -->
</HEAD>
<BODY leftMargin=2 topMargin=2 marginwidth="2" marginheight="2" text="#FF0000" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<font face="Arial" size="1" color="#FF0000">
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
<!---- Main Content is inserted here ------>         
            
<table border="0" cellspacing="0" cellpadding="0" width="100%">
 
  
  <tr> 
    <td width="644" valign="top">  
     <b><a href="DatAdmin.asp"><center><font face="Arial, Helvetica, sans-serif" size="2" color="#FF0000">Back to 
      Administrator Management</font></center> 
      </a></b>  <br>
      <div align="center"><b>  
        <p>  
          <font color="#FF0000"> 
          <%
		  pname=trim(request("pname"))
		  if pname="" then
				strsql="SELECT product.description, product.new, product.Model, exportprice.Exportprice, category.Category, " _
				      &"product.pic2path, product.picpath,product.id " _
					  &"FROM category RIGHT JOIN (exportprice RIGHT JOIN " _
					  &"product ON exportprice.Id = product.exportprice) " _
					  &" ON category.Id = product.category "
          else
				strsql="SELECT product.description, product.new, product.Model, exportprice.Exportprice, category.Category, " _
				      &"product.pic2path, product.picpath,product.id " _
					  &"FROM category RIGHT JOIN (exportprice RIGHT JOIN " _
					  &"product ON exportprice.Id = product.exportprice) " _
					  &" ON category.Id = product.category where product.model like '%"&pname&"%' "
		  end if
				set acres=nothing 
				set acres=createobject("adodb.recordset") 
				acres.cursortype=1 
				acres.locktype=3 
				acres.open strsql,conn 
				recount=acres.recordcount 
				if pageno="" then 
				 pageno=1 
				end if 
				if pageno>1 then 
				  i=(pageno-1)*10 
				  acres.move i 
				 end if  
				if not acres.eof then  
			%>
          </font>
        </p>
        <form name="editproduct" action="sa_new.asp" method="post">  
              <center>
            <font color="#FF0000">
            <input type="text" name="pname" maxlength="30" size="18" value="<%=pname%>">
            <input type="button" value="Find Model" onclick="seachmodel();"> 
            <input type="submit" name="Submit2" value=" Add selected product to Special offer "></font></center> 

          <table width="91%" border="1" cellspacing="2" cellpadding="0" height="45"> 
            <tr>  
              <td width="5%">  
                <div align="center"><b><font color="#FF0000" face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font></b></div> 
              </td> 
              <td width="11%">  
                <div align="center"><b><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000">Model</font></b></div> 
              </td> 
              <td width="27%">  
                <div align="center"><b><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000">Description</font></b></div> 
              </td> 
              <td width="20%">  
                <div align="center"><b><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000">Small  
                  Photo</font></b></div> 
              </td> 
              <td width="20%">  
                <div align="center"><font size="2"><b><font face="Arial, Helvetica, sans-serif" color="#FF0000">Big  
                  Photo</font></b></font></div> 
              </td> 
              <td width="17%">  
                <div align="center"><b><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000">Price  
                  Range</font></b></div> 
              </td> 
            </tr> 
            <%    
			  	do while not acres.eof   
				   if j>9 then exit do  
			  %> 
            <tr>  
              <td width="5%" height="94">  
<% 
new = acres("new")
'response.write new
%>
                <div align="center"> <font face="Arial, Helvetica, sans-serif" color="#FF0000">  
                  <input type="checkbox" <% if new=1 then %> checked <% end if %> name="product_id" value="<%= trim(acres("id")&"") %>" > 
                  </font></div> 
              </td> 
<input type=hidden value="<%= trim(acres("id")&"") %>" name="zero"> 
              <td width="11%" height="94"><font color="#FF0000"><font size="2" face="Arial, Helvetica, sans-serif"><%= trim(acres("model")&"") %></font><font face="Arial, Helvetica, sans-serif">&nbsp;</font></font></td>
              <td width="27%" height="94"> 
                <div align="center"><font color="#FF0000"><font size="2" face="Arial, Helvetica, sans-serif"><%= trim(acres("description")&"") %></font><font face="Arial, Helvetica, sans-serif">&nbsp;</font></font></div>
              </td>
              <td width="20%" height="94" valign="middle"> 
                <div align="center"><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000"> 
                  <%if not isnull(acres("picpath")) then%>
				  <a href="showimg.asp?id=<%=acres("id")%>&flage=1" target="_blank" title="show bigger">
                  <img src="showimg.asp?id=<%=acres("id")%>&flage=1" border=1 width="100" height="80">
				  </a>
                  <%else
                      response.write "no photo"
                    end if%>
                  </font></div>
              </td>
              <td width="20%" height="94" valign="middle"> 
                <div align="center"><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000"> 
                  <%if not isnull(acres("pic2path")) then%>
				  <a href="showimg.asp?id=<%=acres("id")%>&flage=2" target="_blank" title="show bigger">
                  <img src="showimg.asp?id=<%=acres("id")%>&flage=2" border=1 width="100" height="80">
				  </a>
                  <%else
                      response.write "no photo"
                    end if%>
                  </font></div>
              </td>
              <td width="17%" height="94"> 
                <div align="center"><font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000"><%= trim(acres("exportprice")&"") %>&nbsp;</font></div>
              </td>
            </tr>
            <%  
			  		acres.movenext   
					j=j+1 
			    loop  
			 call showpageno(pageno) %>
          </table>  
			  
          <table width="100%">
            <tr> 
              <td width="3%"><font color="#FF0000">&nbsp;</font> </td>
              <td width="97%"> 
                <font color="#FF0000"> 
                <%call showpageno(pageno)%>
                </font>
              </td>
            </tr>
          </table> 
			</form>  
        <font color="#FF0000">  
			<% Else  %>  
				No Records! 
			<% End If %>   
	   
        </font>   
	   
	  </b></div> 
    </td> 
  </tr> 
</table> 
 
<!---- Main Content is ended here ------>   
 
</TD> 
<!--#include file="menu_bottom.inc"--> 
</BODY> 
 
<%   
function showpageno(pageno)  
	if recount>10 then  
		lastpage=recount\10  
		if recount mod 10 >0 then  
			lastpage=lastpage+1  
		end if  
		if pageno>10 then  
		     response.write "<a href='product_list.asp?pageno=1&pname="&pname&"'> The First Page </a>&nbsp;&nbsp;"  
			response.write "<a href='product_list.asp?pageno="&(pageno-9-(pageno  mod 10) )&"&pname="&pname&"'>Previous 10</a>&nbsp;&nbsp;"  
		end if  
		strtemp=pageno  
		if (pageno Mod 10 )=0 then  
		   strtemp=strtemp-10  
		end if  
	 for i=(strtemp-(strtemp mod 10)+1) to (strtemp+10-(strtemp mod 10))  
	         if lastpage<i then  exit for			   
            if i- pageno=0 then  
				response.write cstr(i)&"&nbsp;&nbsp;"  
			else  
				response.write "<a href='product_list.asp?Pageno="&i&"&pname="&pname&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;"  
			end if	  
		next  
		if (pageno\10)<(lastpage\10) then  
		        response.write "<a href='product_list.asp?Pageno="&(pageno+11-(pageno mod 10))&"&pname="&pname&"'>Next 10</a>&nbsp;&nbsp;"  
			   response.write "<a href='product_list.asp?Pageno="& (lastpage) &"&pname="&pname&"'> The Last Page</a>&nbsp;&nbsp;"  
		end if  
		  
 end if  
end function  
%>