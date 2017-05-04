<!--#include file="database.inc"-->
<%
strpid=request("id")
isfind=false
if trim(strpid)<>"" then
	strsqldet="SELECT product.description,product.ID, product.Model, category.Category,  " _
			 &"product.moneytype,product.Pic2Path, " _
			 &"product.packing, product.minorder, product.delivery, exportprice.exportprice, " _
			 &"product.term, product.sampleprice " _
			 &"FROM (product LEFT JOIN category ON product.category = category.Id) " _
			 &"LEFT JOIN exportprice ON product.exportprice = exportprice.Id " _
			 &"Where product.id="& strpid

	set acres=conn.execute(strsqldet)
	if not acres.eof then
		isfind=true
		strmodel=trim(acres("model")&"")
		strcategory=trim(acres("category")&"")
		strdes=trim(acres("description"))
		strpic=acres("pic2Path")
		strpak=trim(acres("packing")&"")
		strorder=trim(acres("minorder")&"")
		strdelivery=trim(acres("delivery")&"")		
		strexprice=trim(acres("exportprice")&"")				
		strterm=trim(acres("term")&"")						
		strsampleprice=trim(acres("sampleprice")&"")
		strmoneytype=trim(acres("moneytype")&"")
	
	end if 
end if
%>
<html>
<head>
<title>detail</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK title=style1 href="images/basic.css" type=text/css rel=stylesheet><!--BEGIN menuItemData.asp -->
<script language="JavaScript">
<!--

function closewin(what)
{window.close(what);
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" bgcolor="#808080" cellspacing="1" cellpadding="2">
  <tr bgcolor="#FFFFFF"> 
    <td width="18%" 
 height="19"> 
      型號:</td>
    <td 
 height="19"> 
      <div align="center">
<%=strmodel%>
</div>
    </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td width="18%"> 
     類型:
    </td>
    <td> 
      <div align="center">
<%=strcategory%>
</div>
    </td>
  </tr>
 <tr bgcolor="#FFFFFF"> 
    <td width="18%"> 
    內容:</td>
    <td> 
      <div align="center">
<%=strdes%>
</div>
    </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td width="18%"> 
      參考資料: </td>
    <td align="center"> 
<% If strpak <> "" Then %>  
   <a href="<% = strpak %>">按此下戴</a>
        <% Else %>
               暫末提供
               <% End If%>


    </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td width="18%" 
> 
      <div align="left">價錢:

    </td>
    <td 
> 
      <% if strexprice="" then%>
      <div align="center">
 聯絡我們</div>
      <%else%>
      <div align="center">
<%=strexprice%>
</div>
      <%end if%>
    </td>
  </tr>

  <input type="hidden" name="pcid" value="">
</table>
<div align="center">
  <input type="button" name="Button" value="Close Window" onClick="closewin('detail.asp')">
</div>
</body>
</html>