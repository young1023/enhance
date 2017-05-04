<!--#include file="database.inc"-->
<% 
response.expires=0 
straction=trim(request("action_button"))
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
if straction<>"" then
	strsearchsql="SELECT product.description, product.Model, exportprice.Exportprice, category.Category, " _
	      &"product.pic2path, product.id " _
		  &"FROM category RIGHT JOIN (exportprice RIGHT JOIN " _
		  &"product ON exportprice.Id = product.exportprice) " _
		  &" ON category.Id = product.category " _
		  &"where product.new=1 "
end if
if lcase(straction)="subcategory" then
	strsearchsql="SELECT product.description, product.Model, exportprice.Exportprice, category.Category, " _
	      &"product.pic2path, product.id " _
		  &"FROM category RIGHT JOIN (exportprice RIGHT JOIN " _
		  &"product ON exportprice.Id = product.exportprice)  " _
		  &"ON category.Id = product.subcategory " _
		  &"where 1=1 "
end if
select case lcase(straction)
	case "model"
		strmodel=trim(request("txtmodel"))
		if strmodel<>"" then
			strmodel=replace(strmodel,"'","''")
			strwhere=" and product.model like '"& strmodel & "%' "
		end if
	case "exprice"
		strexprice=trim(request("selexprice"))
		if strexprice<>"" then
			strwhere=strwhere&" and exportprice.id="& strexprice &" "
		end if
		
	case "category"
	    strcat=trim(request("selcategory"))
		if strcat<>"" then
			strwhere=strwhere & "  and category.id="& strcat &" "
		end if
	case "subcategory"
	    strcat=replace(trim(request("subcategory")),"'","''")
		if strcat<>"" then
			strwhere=strwhere & "  and category.category='"& strcat &"' "
		end if
end select
if session("where")="" then
  strsearchsql=strsearchsql & strwhere
  session("where")=strwhere
 'response.write session("where")
else 
  strwhere=session("where")
  strsearchsql=strsearchsql & strwhere

  'response.write strsearchsql

end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>What's New</title>
<style type="text/css">
 ul {font-family: "Arial", "Helvetica", "sans-serif"; font-size: 11px; color: #000066; font-weight: bold}
 table {  font-family: "Arial", "Helvetica", "sans-serif"; color: #0000FF; font-size: 11px}
 picposition {  margin-left: 15px; margin-top: 0px; margin-bottom: 0px}</style>
<base target="_self">
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

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
       {strmsg="Please input Number into QTY field.\n";
         break;
       }
     if((document.form1.txtt(i).value)<=0)
       {strmsg="please enter a quantity of one or more!.\n";
		break;
       }
     if (document.form1.txtt(i).value.indexOf(".")!=-1)
       {strmsg="please enter a quantity of one or more and must be integer!.\n";
        break;
       }
     }
  }

if(i>0 && k==document.form1.txtt.length)
  strmsg="Please input Number into QTY field.\n";
  //if(i>0&& j==document.form1.txtt.length)
  //strmsg="Please enter a quantity of one or more!.\n";
if(i==0 && j==0)
{
 if (document.form1.txtt.value=="")			    
  strmsg="Please input Number into QTY field.\n";
 else
  { 
   if(isNaN(document.form1.txtt.value))
     strmsg="Please input Number into QTY field.\n";					  
   if((document.form1.txtt.value)<=0)					 
	 strmsg="please enter a quantity of one or more!.\n";			   
   if (document.form1.txtt.value.indexOf(".")!=-1)			 
     strmsg="please enter a quantity of one or more and must be integer!.\n";
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
</script>
<link rel="stylesheet" href="images/Tabcss.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="images/demo/bdg1.jpg" link="ffffff" vlink="ffffff">

<!--#include file="top.inc"-->
<div align="left">
  <table border="0" cellspacing="0" cellpadding="0" width="780">
    <tr> 
      <td valign="top" width="124"> 
        <div align="center"> 
          <!--#include file="Left.inc" -->
        </div>
      </td>
      <td valign="top" align="center" width="666"> 
        <div align="center"> 
          <center>
           
			<% If strsearchsql<>"" then %>
            <p><font size="3"><b><font color="#FFFFCC" face="Arial, Helvetica, sans-serif">What's New</font></b></font></p>
				<font color="ffffff" face="Arial, Helvetica, sans-serif" size=3><b>
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
						i=(pageno-1)*10
						acres.move i
					end if
					if not acres.eof then
				%></b></font>
				 
            <form name="form1" action="enquiry.asp" method="post">
            <table width="95%" border="0">
              <tr  > 
                <td colspan="10"> 
                  <div align="left"> <font color="White" face="Arial" size="2"> 
                    <%call showpageno(pageno,straction,strwhere) 
					  oldstrwhere=strwhere
					 %>
                    </font> </div>
                </td>
              </tr>
              
              <% 	  	do while (not acres.eof)
						if j>9 then exit do
				  %>
              <tr  > 
<td align="center" width="25%" valign="middle"> 
                    <a href="productdetail.asp?id=<%=trim(acres("id")&"") %>"> 
                  <%if not isnull(acres("pic2path")) then%>
                  <img src="showimg.asp?id=<%=acres("id")%>&flage=2" border=1 width="100" height="80">
                  <%else
                      response.write "<font size=2>No Photo</font>"
                    end if%>
                    </a> </td>
             
            <td width="50%" align="left"><font size="2" color="ffffff"><a href="productdetail.asp?id=<%=trim(acres("id")&"")%>">
				  <%=acres("model")%></a></font><a href="productdetail.asp?id=<%=trim(acres("id")&"")%>"></a><br><font size="2" color="ffffcc">
<%=trim(acres("description")&"") %></font>&nbsp;<br>
 <%if trim(acres("exportprice")&"")="" then %>
                    <font size="2"color="ffffcc"> Contact Us </font> 
                    <% Else  %>
                    <font size="2"color="ffffcc"> <%=trim(acres("exportprice")&"") %></font> 
                    <% End If %>
</td>
                  
                  <td width="20%" align="center"> 
                    <div align="center"> 
                      <p><font color="#FFFFCC" size="2" face="Arial, Helvetica, sans-serif">Interest 
                        QTY</font> 
                        <input type="text" name="txtt" size="5" maxlength="12">
                        <input type="hidden" name="txtpoid" value="<%= trim(Acres("id")&"") %>">
                      </p>
                      </div>
                </td>
              </tr>
              <%  		acres.movenext
						j=j+1 
				    loop
				  %>
              <tr > 
                <td colspan="10"> 
                      <div align="center"> 
                        <input type="button" value="Add to Request List" onClick="javascript:dosubmit();" name="button">
                        <%
						 if trim(session("enquiryno"))<>"" then %>
                        <input type="button" name="Button" value="View List" onClick="javascript:location='enquiry.asp'">
						  <%
				          strsql="Select * from Acclist where A_id= "& session("enquiryno") &"  "
				           set acresorder=conn.execute(strsql)				
				          if not Acresorder.eof then %>
				         
                        <input type="button" name="Button2" value="Submit" onClick="javascript:location='sendto.asp'">
                        <input type="button" name="Button3" value="Empty Basket" onClick="javascript:location='deletecar.asp'">
                        <%end if
					       set acresorder=nothing
					      set strsql=nothing
					    %>
                        <input type="hidden" name="hidbuttion" value="">
                        <%end if%>
                      </div>
                      <div align="right"> 
                      <font color="White" face="Arial" size="2"> 
                      <%call showpageno(pageno,straction,oldstrwhere) %>
                      </font> </div>
                </td>
              </tr>
            </table> 
            </form>				
            <% Else  %>
            <font color="#FFFFCC" size="2">No Record Find</font><font color="#FFFFCC">.</font> 
            <% End If %>
            <% End If %>
            <p>&nbsp;</p> 
          </center> 
        </div> 
      </td>      
    </tr> 
  </table> 
  <div align="center"><font size="2" color="#CCCCCC">Copyright&copy; 2001 Winson 
    Development Company All right reserved.</font></div>
</div>  
</body>  
</html> 
<% 
function showpageno(pageno,straction,strwhere) 
	if recount>10 then 
		lastpage=recount\10 
		if recount mod 10 >0 then 
			lastpage=lastpage+1 
		end if 
		if pageno>10 then 
		     response.write "<a href='new.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&pageno=1'> The First Page 10</a>&nbsp;&nbsp;" 
			response.write "<a href='new.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&pageno="&(pageno-9-(pageno  mod 10) )&"'>Previous 10</a>&nbsp;&nbsp;" 
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
				response.write "<a href='new.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;" 
			end if	 
		next 
		if (pageno\10)<(lastpage\10) then 
		       	response.write "<a href='new.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&Pageno="&(pageno+11-(pageno mod 10)) &"'>Next 10</a>&nbsp;&nbsp;" 
			   response.write "<a href='new.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&Pageno="& (lastpage) &"'> The Last Page 10</a>&nbsp;&nbsp;" 
		end if 
		 
 end if 
end function 
%> 
