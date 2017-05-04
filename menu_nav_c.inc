<script language="JavaScript">
<!--
function dosearch(p_arg)
{
	var strmsg="";
	
	switch(p_arg){
		case "model":
			if(document.search.txtmodel.value=="")
				strmsg="You must input a word to search.\n";
			break;
		case "category":
			if(document.search.selcategory.value=="-1")
			strmsg="You must select category to search!\n";
			break;
		case "exprice":
			if(document.search.selexprice.value=="-1" )
			strmsg="You must select a export price to search!\n";
			break;
		default:
	}
	if(strmsg!="")
	alert(strmsg);
	else
	{
		document.search.action_button.value=p_arg;
		document.search.submit();
	}
}
function dologin()
{
	if(document.mainlogin.txtemail.value=="" || document.mainlogin.txtpassword.value=="" )
		alert("Please input login email address and password!");
	else
		document.mainlogin.submit();
}

//-->
</script>
  <form name="search" action="searchresult_c.asp" method="post">
<!-- non-flash table width="760" -->
<TABLE cellSpacing=0 cellPadding=0 width=760 border=0>
  <TBODY>
  <TR height=10>
    <TD bgColor=#e41f1f><IMG height=1 alt="" 
      src="images/space.gif" width=1 border=0></TD>
    <TD bgColor=#e41f1f><IMG height=1 alt="" 
      src="images/space.gif" width=758 border=0></TD>
    <TD bgColor=#e41f1f><IMG height=1 alt="" 
      src="images/space.gif" width=1 border=0></TD></TR>
  <TR>
    <TD bgColor=#e41f1f><IMG height=1 alt="" 
      src="images/space.gif" width=1 border=0></TD>
    <TD bgColor=#e41f1f>
      <TABLE cellSpacing=0 cellPadding=0 width="96%" border="0">
        <TBODY>
        
        <TR>
          <TD bgColor=#e41f1f rowSpan=2><IMG height=45 alt="" 
            src="images/space.gif" width=1 border=0></TD>
          <TD bgColor=#e41f1f rowSpan=2 width=220>&nbsp;&nbsp;&nbsp;</TD> 
          <TD bgColor=#e41f1f colSpan=4>
            <p align="right">&nbsp;&nbsp;&nbsp; <font color="#FFFFFF">Enhance Technologies Limited</font></p>                                               
          </TD>
        </TR>
        <TR>
          <TD bgColor=#e41f1f align="right"><a href="mailto:info@enhance-tech.com.hk"><font color="#FFFFFF">info@enhance-tech.com.hk</a>  |&nbsp;&nbsp;&nbsp;&nbsp;Tel: 852 2304 1036</font></TD>
        
        </TR>
        </TBODY></TABLE></TD>
    <TD bgColor=#e41f1f><IMG height=1 alt="" 
      src="images/space.gif" width=1 border=0></TD></TR>
  <TR>
    <TD bgColor=#e41f1f colSpan=3><IMG height=1 alt="" 
      src="images/space.gif" width=1 border=0></TD></TR>
  <TR>
    <TD bgColor=#e41f1f><IMG height=1 alt="" 
      src="images/space.gif" width=1 border=0></TD>
    <TD bgColor=#ffffff><!-- top navigation -->
      <TABLE cellSpacing=0 cellPadding=0 width=100% border=0>
        <TBODY>
        <TR>
          <TD colSpan=3><IMG height=5 alt="" 
            src="images/space.gif" width=5 border=0></TD></TR>
        <tr>
	<td><img src="/assets/images/space.gif" border="0" width="5" height="1" alt=""></td>
	<td bgcolor="#FFFFFF">
<img src="/assets/images/space.gif" border="0" width="20" height="1" alt=""><a href="default_c.asp"><font color="Black">主頁</font></a>
<img src="/assets/images/space.gif" border="0" width="10" height="1" alt="">|<img src="/assets/images/space.gif" border="0" width="10" height="1" alt="">
<a href="profile_c.asp"><font color="Black">公司資料</font></a><img src="/assets/images/space.gif" border="0" width="10" height="1" alt="">|<img src="/assets/images/space.gif" border="0" width="290" height="1" alt="">|<img src="/assets/images/space.gif" border="0" width="10" height="1" alt="">

<img src="/assets/images/space.gif" border="0" width="5" height="1" alt="">|<img src="/assets/images/space.gif" border="0" width="5" height="1" alt="">

<img src="/assets/images/space.gif" border="0" width="5" height="1" alt="">|<img src="/assets/images/space.gif" border="0" width="5" height="1" alt="">
<a href="contact_c.asp"><font color="Black">聯絡我們</font></a>
<img src="/assets/images/space.gif" border="0" width="5" height="1" alt="">|<img src="/assets/images/space.gif" border="0" width="5" height="1" alt="">
<a href="login.asp"><font color="Black">登入</font></a>
<img src="/assets/images/space.gif" border="0" width="5" height="1" alt="">|<img src="/assets/images/space.gif" border="0" width="5" height="1" alt="">
<a href="default.asp"><font color="Black">English</font></a></td>
	<td rowspan="3"></td>
	</tr>
        <TR>
          <TD><IMG height=1 alt="" src="images/space.gif" 
            width=5 border=0></TD>
          <TD bgColor=#e41f1f><IMG height=1 alt="" 
            src="images/space.gif" width=1 border=0></TD></TR>

        <TR>
          <TD height="20"><IMG  alt="" height=20 src="images/space.gif" 
            width=5 border=0></TD>
          <TD vAlign=bottom bgColor=#ffffff><IMG  alt="" 
            src="images/space.gif" height=20 width=10 border=0><A 
            onmouseover="showMenu('multifunction');window.status='';return true" 
            style="CURSOR: default" onmouseout="hideMenu('multifunction')" 
            >多功能筆</A><IMG  alt="" 
            src="images/space.gif" height=1 width=12 border=0>|<IMG  alt="" 
            src="images/space.gif" height=20 width=12 border=0><A 
            onmouseover="showMenu('metal');window.status='';return true" 
            style="CURSOR: default" onmouseout="hideMenu('metal')" 
            >金屬筆</A><IMG  alt="" 
            src="images/space.gif" height=1 width=12 border=0>|<IMG  alt="" 
            src="images/space.gif" height=20 width=12 border=0><A 
            onmouseover="showMenu('plastic');window.status='';return true" 
            style="CURSOR: default" onmouseout="hideMenu('plastic')" 
            > 塑膠筆 </A><IMG  alt="" 
            src="images/space.gif" height=4 width=12 border=0>|<IMG  alt="" 
            src="images/space.gif" height=20 width=12 border=0><A 
            onmouseover="showMenu('wooden');window.status='';return true" 
            style="CURSOR: default" onmouseout="hideMenu('wooden')" 
            > 木制筆</A><IMG  alt="" 
            src="images/space.gif" height=4 width=12 border=0>|<IMG  alt="" 
            src="images/space.gif" height=20 width=12 border=0><A 
            onmouseover="showMenu('led');window.status='';return true" 
            style="CURSOR: default" onmouseout="hideMenu('led')" 
            > LED系列</A><IMG  alt="" 
            src="images/space.gif" height=4 width=12 border=0>|<IMG  alt="" 
            src="images/space.gif" height=20 width=12 border=0><A onmouseover="showMenu('packaging');window.status='';return true" onmouseout="hideMenu('packaging')" style="cursor: default;">
包裝盒
</A><IMG  alt="" 
            src="images/space.gif" height=4 width=12 border=0>|<IMG  alt="" 
            src="images/space.gif" height=20 width=12 border=0><a href="searchresult_c.asp?selcategory=47&action_button=category">雷射系列</a><IMG  alt="" 
            src="images/space.gif" height=4 width=10 border=0>|<IMG  alt="" 
            src="images/space.gif" height=20 width=12 border=0><a href="searchresult_c.asp?selcategory=59&action_button=category">其他</a><IMG  alt="" 
            src="images/space.gif" height=4 width=10 border=0>|<IMG  alt="" 
            src="images/space.gif" height=20 width=10 border=0>
<A 
            onmouseover="showMenu('price');window.status='';return true" 
            style="CURSOR: default" onmouseout="hideMenu('price')" 
            ><font color="Black">價格搜索</font></A><IMG  alt="" 
            src="images/space.gif" height=4 width=12 border=0>|<IMG  alt="" 
            src="images/space.gif" height=20 width=12 border=0> <input type="text" name="txtmodel" size="10" style=" FONT-SIZE: 7pt" value="" maxlength="10">   
              <input type="submit" name="action_button" width="6" height="6" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 8pt ; FONT-color: red" value="Model">
</TD></TR>
</form>
 <TR>
          <TD><IMG height=1 alt="" src="images/space.gif" 
            width=5 border=0></TD>
          <TD>

</TD></TR>
</TBODY></TABLE></TD>
    <TD bgColor=#e41f1f><IMG height=1 alt="" 
      src="images/space.gif" width=1 border=0></TD></TR>
  
  <TR>
    <TD bgColor=#e41f1f><IMG height=10 alt="" 
      src="images/space.gif" width=1 
border=0><!-- here --></TD>
    <TD vAlign=top align=middle bgColor=#ffffff>
