<%
if session(("SECURITY_LEVEL"))<>"1" then
  response.redirect "default.asp"
end if

id=trim(request("id"))
if id="" then
  response.redirect "sa_tradeshow.asp"
end if
%>
<!--#include file="included/conne.inc" -->
<%
sql="select * from tradeshow where id="&id&" "
set rs=conn.execute(sql)
if not rs.eof then
  trade_name=trim(rs("trade_name"))
  location=trim(rs("location"))
  trade_link=trim(rs("http"))
  trade_date=trim(rs("trade_date"))
  trade_id=trim(rs("trade_id"))
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<HTML><HEAD><TITLE>Edit Tradeshow</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="included/default.css" type="text/css">
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=4>
<DIV align=center>
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
            <TD height=18 width="180" align=center><script language=JavaScript src="included/date.js"></script>
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
            <td><included src="included/back.gif" width="306" height="37"></td>
          </tr>
        </table>
        <table width="100%" border="0" cellpadding="1" cellspacing="1" height="100%">
          <tr> 
            <td bgcolor="#000000">
              <table width="100%" border="0" cellpadding=0 cellspacing="0" bgcolor="#FFFFFF" height="100%">
                <tr>
                  <td valign="top" align="center" bgcolor="#FFCC33" width="180">

                  </td>
                  <td valign="top" align="center" bgcolor="#E6EBEF">
                    <table width="93%" border="1" cellpadding="0" cellspacing="0" bgcolor="#E6EBEF">
                      <form name=fm1 method=post>
                        <tr> 
                          <td height="48" align="center" colspan="2"><b><font color="#FF6600">Edit Tradeshow Information</font></b> 
                            <SCRIPT language=JavaScript>
<!--
function check()
{
document.fm1.action="execute.asp";
if(document.fm1.title.value=="")
 {
  alert("Please Input Trade name");
  document.fm1.title.focus();
  return false;
 }
if(document.fm1.location.value=="")
 {
  alert("Please Input Location");
  document.fm1.location.focus();
  return false;
 }

document.fm1.submit();
}
//-->
</SCRIPT>
                          </td>
                        </tr>
 <tr> 
                          <td align="right" height="28" width="50%">Trade ID : &nbsp;</td>
                          <td align="left">
                            <input type="text" name="typeid" maxlength="30" size="23" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #FFCC33 1px solid; BORDER-LEFT: #FFCC33 1px solid; BORDER-RIGHT: #FFCC33 1px solid; BORDER-TOP: #FFCC33 1px solid; FONT-SIZE: 10pt" value="<%=trade_id%>">
                            </td>
                        </tr>
                        <tr> 
                          <td align="right" height="28" width="50%">Trade Name : &nbsp;</td>
                          <td align="left">
                            <input type="text" name="title" maxlength="30" size="40" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #FFCC33 1px solid; BORDER-LEFT: #FFCC33 1px solid; BORDER-RIGHT: #FFCC33 1px solid; BORDER-TOP: #FFCC33 1px solid; FONT-SIZE: 10pt" value="<%=trade_name%>">
                            </td>
                        </tr>
                        <tr> 
                          <td align="right" height="28" width="50%">Trade Location : &nbsp;
                          </td> <td align="left">
                          <input type="text" name="location" maxlength="50" size="40" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #FFCC33 1px solid; BORDER-LEFT: #FFCC33 1px solid; BORDER-RIGHT: #FFCC33 1px solid; BORDER-TOP: #FFCC33 1px solid; FONT-SIZE: 10pt" value="<%=location%>">  </td>
                        </tr>
<tr> 
                          <td align="right" height="28" width="50%">Hyperlink : &nbsp;
                           </td> <td align="left">
                          <input type="text" name="trade_link" maxlength="100" size="40" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #FFCC33 1px solid; BORDER-LEFT: #FFCC33 1px solid; BORDER-RIGHT: #FFCC33 1px solid; BORDER-TOP: #FFCC33 1px solid; FONT-SIZE: 10pt" value="<%=trade_link%>"></td>
                        </tr>
                        <tr> 
                          <td height="28" align="right">Date : &nbsp;</td>
                           <td align="left"> <input type="text" name="postdate" maxlength="50" size="23" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #FFCC33 1px solid; BORDER-LEFT: #FFCC33 1px solid; BORDER-RIGHT: #FFCC33 1px solid; BORDER-TOP: #FFCC33 1px solid; FONT-SIZE: 10pt" value="<%=trade_date%>">
                         </td>
                        </tr>
                        <tr align="center"> 
                          <td height="48" colspan="2"> 
                            <input type="button" value="Update" onclick="check();" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 10pt; HEIGHT: 20px; WIDTH: 80px">
						  <%if request("yn")=1 then%>
                          <input type="button" value="Return" onClick="history.go(-3);" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 9pt; HEIGHT: 20px; WIDTH: 80px">
						  <%else%>
                          <input type="button" value="Return" onClick="history.go(-1);" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 9pt; HEIGHT: 20px; WIDTH: 80px">
						  <%end if%>
                            <input type="reset" value="Reset" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 10pt; HEIGHT: 20px; WIDTH: 80px">
                            <input type=hidden value="mtradeshow" name=whatdo>
                            <input type=hidden value="<%=id%>" name=id>
                          </td>
                        </tr>
                      </form>
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
        <script language=JavaScript src="included/copyright.js"></script>
      </TD>
    </TR>
</TABLE>
</DIV>
</BODY></HTML>