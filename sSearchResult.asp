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
	strsearchsql="SELECT product.description,product.cdescription, product.Model, exportprice.Exportprice, category.Category, " _
	      &"product.pic2path, product.id " _
		  &"FROM category RIGHT JOIN (exportprice RIGHT JOIN " _
		  &"product ON exportprice.Id = product.exportprice) " _
		  &" ON category.Id = product.category " _
		  &"where 1=1 "
end if
if lcase(straction)="subcategory" then
	strsearchsql="SELECT product.description,product.cdescription, product.Model, exportprice.Exportprice, category.Category, " _
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
<meta http-equiv="Content-Type" content="text/html; charset=GB2312">
<title>Search Result</title>
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
<p align="center"><img src="images/layout-header_c.jpg"></p>
<div align="left">
  <table border="0" cellspacing="0" cellpadding="0" width="780">
    <tr> 
      <td valign="top" width="124"> 
        <div align="center"> 
          <!--#include file="sLeft.inc" -->
        </div>
      </td>
      <td valign="top" align="center" width="666"> 
        <div align="center"> 
          <center>
            <!--#include file="ssbar.inc" -->
			<% If strsearchsql<>"" then %>
             <p><font size="2" color="#EFFBBF">搜索结果</p>
          <b>
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
				<% if recount=1 then%>
				  <p><font color="#FFFFcc" size="2" face="Arial, Helvetica, sans-serif"><b> 
        找到</b></font>
		   <font color="ffffff" size=3><b><%=recount%></b></font>
		  <font size="2" color="#EFFBBF">条记录</font></p>			           			           
                <%else%>
			<p><font size="2" color="#EFFBBF">找到<font color="#000000"> 
		   <font color="ffffff" size=3><b><%=recount%></b></font> 
                <font size="2" color="#EFFBBF">条记录</font></p>
            <font color="#000000">	
		    <%end if %>
            <font color="#FFFFcc" size="2">请输入您感兴趣的产品数量，并且按 "加入要求报价表 " 按钮<br>想了解产品详情。可按"型号"或"图片"的超连结
           <br>
            <br>
            <form name="form1" action="senquiry.asp" method="post">
            <table width="83%" border="1">
              <tr  > 
                <td colspan="10"> 
                  <div align="left"> <font color="White" face="Arial" size="2"> 
                    <%call showpageno(pageno,straction,strwhere) 
					  oldstrwhere=strwhere
					 %>
                    </font> </div>
                </td>
              </tr>
              <tr > 
                  <td width="17%"> 
                    <div align="center"><font size="3"><b><font color="#FFFFcc"> 
                      <font face="Arial, Helvetica, sans-serif" size="2"> 
                      <% If instr(1,strwhere,"order by")>0 then
						strwhere=mid(strwhere,1,instr(1,strwhere,"order")-1)
					end if
				  		strwhere=strwhere & " order by product.Model"
				  %>
                       <a href="sSearchresult.asp?action_button=<%= server.URLencode(straction) %>&strwhere=<%= server.URLencode(strwhere) %>&pageno=<%= 1 %>"> 
                      <img src="images/product/up.gif" width="15" height="15" border="0"></a> <font color="#FFFFcc">型号</font>
                      <% If instr(1,strwhere,"order by")>0 then
						strwhere=mid(strwhere,1,instr(1,strwhere,"order")-1)
					end if
				  		strwhere=strwhere & " order by product.Model DESC "
				  %>
                      <a href="sSearchresult.asp?action_button=<%= server.URLencode(straction) %>&strwhere=<%= server.URLencode(strwhere) %>&pageno=<%= 1 %>"> 
                      <font color="#FFFFcc"><img src="images/product/down.gif" width="15" height="15" border="0"></font></a> 
                      </font></font></b></font></div>
                </td>
                    <td width="22%"> 
                    <div align="center"><font size="3"><b><font color="#EFFBBF" size="2">内容</font></b></font></div>
                </td>
                  <td width="25%"> 
                    <div align="center"><font size="3"><b><font color="#EFFBBF" size="2">相片</font></b></font></div>
                </td>
                  <td width="23%"> 
                    <div align="center"><font size="3"><b><font color="#FFFFcc"> 
                      <font face="Arial, Helvetica, sans-serif" size="2"> 
                      <% If instr(1,strwhere,"order by")>0 then
				  		strwhere=mid(strwhere,1,instr(1,strwhere,"order")-1)
					 end if 
					 strwhere=strwhere & " order by exportprice.Exportprice"
				  %>
                       <a href="sSearchresult.asp?action_button=<%= server.URLencode(straction) %>&strwhere=<%= server.URLencode(strwhere) %>&pageno=<%= 1 %>"> 
                      <img src="images/product/up.gif" width="15" height="15" border="0"> <font color="#FFFFcc">价格范围</font>
                      </a> 
                      <% If instr(1,strwhere,"order by")>0 then
				  		strwhere=mid(strwhere,1,instr(1,strwhere,"order")-1)
					 end if 
					 strwhere=strwhere & " order by exportprice.Exportprice desc"
				  %>
                      <a href="sSearchresult.asp?action_button=<%= server.URLencode(straction) %>&strwhere=<%= server.URLencode(strwhere) %>&pageno=<%= 1 %>"><img src="images/product/down.gif" width="15" height="15" border="0"></a>
                      </font></font></b></font></div>
                </td>
                  <td width="13%"> 
                    <div align="center"><font size="3"><b><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFcc">要求报价</font></b></font></div>
                </td>
              </tr>
              <% 	  	do while (not acres.eof)
						if j>9 then exit do
				  %>
              <tr  > 
              <td width="17%" align="center" bgcolor="#003366">
                <font size="2" color="ffffff"><a href="sproductdetail.asp?id=<%=trim(acres("id")&"")%>">
				  <%=acres("model")%></a></font><a href="sproductdetail.asp?id=<%=trim(acres("id")&"")%>"></a></td>
                <!--<td width="8%" align="left"><%=acres("category")%></td> -->
                  <td width="22%" align="left" bgcolor="#003366"><font size="2" color="ffffcc"><%=trim(acres("cdescription")&"") %></font>&nbsp;</td>
                  <td align="center" width="25%" bgcolor="#003366" valign="middle"> 
                    <a href="sproductdetail.asp?id=<%=trim(acres("id")&"") %>"> 
                  <%if not isnull(acres("pic2path")) then%>
                  <img src="showimg.asp?id=<%=acres("id")%>&flage=2" border=1 width="100" height="80">
                  <%else
                      response.write "<font size=2>No Photo</font>"
                    end if%>
                    </a> </td>
                  <td width="23%" align="left" bgcolor="#003366"> 
                    <%if trim(acres("exportprice")&"")="" then %>
                    <font size="2"color="ffffcc">联络我们</font> 
                    <% Else  %>
                    <font size="2"color="ffffcc"> <%=trim(acres("exportprice")&"") %></font> 
                    <% End If %>
                  </td>
                  <td width="13%" align="center" bgcolor="#003366"> 
                    <div align="center"> 
                      <p><font color="#FFFFCC" size="2" face="Arial, Helvetica, sans-serif">感兴趣数量</font> 
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
                        <input type="button" value="加入请求报价表" onClick="javascript:dosubmit();" name="button">
                        <%
						 if trim(session("enquiryno"))<>"" then %>
                        <input type="button" name="Button" value="查看列表" onClick="javascript:location='senquiry.asp'">
						  <%
				          strsql="Select * from Acclist where A_id= "& session("enquiryno") &"  "
				           set acresorder=conn.execute(strsql)				
				          if not Acresorder.eof then %>
				         
                        <input type="button" name="Button2" value="提交" onClick="javascript:location='ssendto.asp'">
                        <input type="button" name="Button3" value="清空列表" onClick="javascript:location='deletecar.asp'">
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
            <font color="#FFFFCC" size="2">没有记录</font><font color="#FFFFCC">.</font> 
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
		     response.write "<a href='sSearchresult.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&pageno=1'> The First Page 10</a>&nbsp;&nbsp;" 
			response.write "<a href='sSearchresult.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&pageno="&(pageno-9-(pageno  mod 10) )&"'>Previous 10</a>&nbsp;&nbsp;" 
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
				response.write "<a href='sSearchresult.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&Pageno="&i&"'>  "&cstr(i)&" </a>&nbsp;&nbsp;" 
			end if	 
		next 
		if (pageno\10)<(lastpage\10) then 
		       	response.write "<a href='sSearchresult.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&Pageno="&(pageno+11-(pageno mod 10)) &"'>Next 10</a>&nbsp;&nbsp;" 
			   response.write "<a href='sSearchresult.asp?action_button="&server.URlencode(straction)&"&strwhere="& server.URLencode(strwhere)&"&Pageno="& (lastpage) &"'> The Last Page 10</a>&nbsp;&nbsp;" 
		end if 
		 
 end if 
end function 
%> 
