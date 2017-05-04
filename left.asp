<%'response.expires=0
    if  session("em")="" or session("id")="" then
        response.redirect "gologin.asp"
    end if
	 open=trim(request("leftopen"))
     if open=""  then
	   open=1
	 end if
	%>
<html>
<head>
<title>Left</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script language="javascript">
<!--
function doleftclose()
{document.frmleft.leftclose.value="close";
document.frmleft.submit();
}
function doleftopen()
{document.frmleft.leftopen.value="open";
document.frmleft.submit();
}
-->
</script>
</head>
<body text="#0000FF" bgcolor="#FFFFCC" alink="#0000FF" link="#0000FF" vlink="#0000ff" leftmargin="0" topmargin="0" >
<p>&nbsp;</p>
<p>&nbsp;</p>
<form name="frmleft" method="post" action="left.asp">
  <table width="148" border="0" cellspacing="0" cellpadding="0">
    <tr bgcolor="#ffffcc"> 
      <td height="31" colspan="2"><img src="image1/emailpage.gif" width="126" height="27"></td>
    </tr>
    <%if open=1 then%>
    <tr bgcolor="#ffffcc"> 
      <td height="8" colspan="2"> 
        <div align="center"><a href="left.asp?leftopen=2"><img src="image1/admin_function.gif" width="126" height="17" border="0" ></a></div>
      </td>
    </tr>
    <%end if
     if open=2 then%>
    <tr bgcolor="#ffffcc"> 
      <td height="5" colspan="2"> 
        <div align="center"><a href="left.asp?leftopen=1"><img src="image1/simple_function.gif" width="126" height="17" border="0" ></a></div>
      </td>
    </tr>

    <tr bgcolor="#CCFFCC"> 
      <td width="21" height="2"> 
        <div align="center"></div>
        　 </td>
      <td width="127" height="2"> <font size="-1"><a href="add.asp" target="mainFrame"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif">Add 
        Member</font></a> </font></td>
    </tr>
    <tr bgcolor="#CCFFCC"> 
      <td width="21" height="6"> 
        <div align="center"></div>
        <p>　</p>
      </td>
      <td width="127" height="6"> <font size="-1" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><a href="edit_member.asp" target="mainFrame">Edit 
        Member</a> </font></td>
    </tr>
    <tr bgcolor="#CCFFCC"> 
      <td height="17" width="21"> 
        <div align="center"></div>
        <font size="2" face="Verdana,Geneva,Arial,Helvetica,sans-serif"></font></td>
      <td height="17" width="127"> <font size="-1" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><a href="delete_member.asp" target="mainFrame">Delete 
        Member</a> </font></td>
    </tr>
    <tr bgcolor="#66FFCC"> 
      <td height="15" width="21"> 
        <div align="center"></div>
        <font size="2" face="Verdana,Geneva,Arial,Helvetica,sans-serif"></font></td>
      <td height="15" width="127"> <font size="-1" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><a href="add_group.asp" target="mainFrame">Add 
        Group</a> </font></td>
    </tr>
    <tr bgcolor="#66FFCC"> 
      <td width="21" height="10"> 
        <div align="center"></div>
        <font size="2" face="Verdana,Geneva,Arial,Helvetica,sans-serif"></font></td>
      <td width="127" height="10"> <font size="-1" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><a href="group_list.asp" target="mainFrame">Edit 
        Group</a> </font></td>
    </tr>
    <tr bgcolor="#66FFCC"> 
      <td height="2" width="21"> 
        <div align="center"></div>
        <font size="2" face="Verdana,Geneva,Arial,Helvetica,sans-serif"></font></td>
      <td height="2" width="127"> <font size="-1" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><a href="delete_group.asp" target="mainFrame">Delete 
        Group</a> </font></td>
    </tr>
    <tr bgcolor="#CCFFCC"> 
      <td height="4" width="21"> 
        <div align="center"></div>
        <font size="2" face="Verdana,Geneva,Arial,Helvetica,sans-serif"></font></td>
      <td height="4" width="127"> <font size="-1" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><a href="add_to_group.asp" target="mainFrame">Add 
        To Group</a> </font></td>
    </tr>
    <tr bgcolor="#CCFFCC"> 
      <td height="11" width="21"> 
        <div align="center"></div>
        　</td>
      <td height="11" width="127"> <font size="-1" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><a href="delete_from_group.asp" target="mainFrame">Delete 
        From Group</a></font></td>
    </tr>
    <tr bgcolor="#66FFCC"> 
      <td height="2" width="21"> 
        <div align="center"></div>
        <font size="2" face="Verdana,Geneva,Arial,Helvetica,sans-serif"></font></td>
      <td height="2" width="127"><font size="-1" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><a href="delete_email.asp" target="mainFrame">Announcement</a></font> 
      </td>
    </tr>
	    
    <tr bgcolor="#66FFCC"> 
      <td width="21" height="6"> 
        <div align="center"></div>
        　</td>
      <td width="127" height="6"> <font face="Verdana, Arial, Helvetica, sans-serif" size="-1"><a href="admin_setup.asp" target="mainFrame">Server 
        Setup</a></font></td>
    </tr>
    <%end if%>
    <tr bgcolor="#ffffcc"> 
      <td height="3" colspan="2"> 
        <div align="center"><font size="2" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><a href="send1.asp" target="mainFrame"><img src="image1/sendemail.gif" width="126" height="17" border="0"></a></font></div>
      </td>
    </tr>
    <tr bgcolor="#ffffcc"> 
      <td height="3" colspan="2"> 
        <div align="center"><font size="2" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><a href="recieve_e_mail.asp" target="mainFrame"><img src="image1/reciev.gif" width="126" height="17" border="0"></a></font></div>
        </td>
    </tr>
    <tr bgcolor="#ffffcc"> 
      <td colspan="2"> 
        <div align="center"><font size="2" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><a href="logout.asp" target="_parent"><img src="image1/logout.gif" width="126" height="17" border="0"></a></font></div>
        <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <input type="hidden" name="leftclose" value="">
        <input type="hidden" name="leftopen" value="">
        </font></td>
    </tr>
  </table> 
  </form> 
</body> 
</html> 
