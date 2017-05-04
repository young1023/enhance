 
<p><a href="http://www.jaguarpen.com"><img src="images/product/bbt-home.gif" alt="Home" border="0" width="103" height="30"></a><br>
  <a href="http://www.jaguarpen.com/new_metal1.htm"><img src="images/whatnew.gif" width="103" height="30" border="0"></a><br>
  <a href="http://www.jaguarpen.com/ppen.htm"><img src="images/new_me8.gif" width="103" height="30" border="0"></a><br>
  <a href="http://www.jaguarpen.com/product.htm"><img src="images/product/bbt-scat.gif" border="0" alt="Product" width="103" height="30"></a><br>
  <a href="http://asp.jaguarpen.com/searchresult.asp"><img src="images/product/bbt-shop.gif" alt="bbt-shop.gif (1252 bytes)" border="0" width="103" height="30"></a><br>

  <a href="contact.asp"><img src="images/product/bbt-contact.gif" alt="bbt-join.gif (1265 bytes)" border="0" width="103" height="30"></a><br>
<a href="http://www.jaguarpen.com/profile.htm"><img src="images/product/bbt-profile.gif" alt="bbt-profile.gif (1068 bytes)" border="0" width="103" height="30"></a><br>
<% if trim(session("user_id"))<>"" then %>  
   <% if trim(session("SECURITY_LEVEL"))="1"then%>
     <a href="DatAdmin.asp"><img src="images/product/adminpage.gif" alt="bbt-profile.gif (1068 bytes)" border="0" width="103" height="30"></a><br>
  <% end if%>  
    <a href="logout.asp"><img src="images/product/bbt-Logout.gif" width="103" height="30" border="0"></a> 
<% End If %>
<br>
<!-- START CODE -->
<FORM ACTION="http://lb.bcentral.com/ex/manage/subscriberprefs.aspx" METHOD="POST">
<FONT SIZE="-1" FACE="arial, helvetica">
<font color="White" size="1">Enter your email and click 'Join Mailing List' button for new products and special offers!</font>
</FONT>
<BR>
<INPUT TYPE="TEXT" NAME="email" />
<INPUT TYPE="HIDDEN" NAME="customerid" value="22180" />
<INPUT TYPE="IMAGE" SRC="images/bbt-join.gif" ALIGN="ABSMIDDLE" BORDER="0" />
</FORM>
<!-- END CODE -->
 <br>
  <a href="info.htm"><b><font color="#EFFBBF" size="2">Terms and Condition</font></b></a> 
</p>
