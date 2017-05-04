 
<p><a href="http://asp.jaguarpen.com/default_c.asp"><img src="images/product/cbbt-home.gif" alt="Home" border="0" width="103" height="30"></a><br>
  <a href="http://www.jaguarpen.com/new_metal1.htm"><img src="images/cwhatnew.gif" width="103" height="30" border="0"></a><br>
  <a href="http://www.jaguarpen.com/product.htm"><img src="images/product/cbbt-scat.gif" border="0" alt="Product" width="103" height="30"></a><br>
  <a href="http://asp.jaguarpen.com/csearchresult.asp"><img src="images/product/cbbt-shop.gif" alt="cbbt-shop.gif (1252 bytes)" border="0" width="103" height="30"></a><br>

  <a href="ccontact.asp"><img src="images/product/cbbt-contact.gif" alt="cbbt-join.gif (1265 bytes)" border="0" width="103" height="30"></a><br>
<a href="http://www.jaguarpen.com/profile.htm"><img src="images/product/cbbt-profile.gif" alt="cbbt-profile.gif (1068 bytes)" border="0" width="103" height="30"></a><br>
<a href="http://www.jaguarpen.com"><img src="images/english.gif" alt="english.gif (1068 bytes)" border="0" width="103" height="30"></a><br>
<% if trim(session("user_id"))<>"" then %>  
   <% if trim(session("SECURITY_LEVEL"))="1"then%>
     <a href="DatAdmin.asp"><img src="images/product/adminpage.gif" alt="cbbt-profile.gif (1068 bytes)" border="0" width="103" height="30"></a><br>
  <% end if%>  
    <a href="clogout.asp"><img src="images/cbbt-Logout.gif" width="103" height="30" border="0"></a> 
<% End If %>
<br>
<!-- START CODE -->
<FORM ACTION="http://lb.bcentral.com/ex/manage/subscriberprefs.aspx" METHOD="POST">
<FONT SIZE="-1" FACE="arial, helvetica">
<font color="yellow" size="2">如想收到最新產品資料,請輸入電郵地址, 然後按 '加入電郵列' 按鈕.</font>
</FONT>
<BR>
<INPUT TYPE="TEXT" NAME="email" />
<INPUT TYPE="HIDDEN" NAME="customerid" value="22180" />
<INPUT TYPE="IMAGE" SRC="images/cbbt-join.gif" ALIGN="ABSMIDDLE" BORDER="0" />
</FORM>
<!-- END CODE -->
 <br>
  <a href="info.htm"><b><font color="#EFFBBF" size="2">Terms and Condition</font></b></a> 
</p>
