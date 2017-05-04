<%
id=trim(request("id"))
if id="" then
 response.redirect "index.asp"
end if
%>
<!--#include file="included/conne.inc" -->
<%
sql="select title,content,createtime from message where id="&id&" "

set rs=conn.execute(sql)
if not rs.eof then
  title=rs("title")
  content=rs("content")
  createtime=rs("createtime")
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<HTML><HEAD><TITLE>News Detail</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="included/home.css" type="text/css">
</HEAD>
<BODY bgColor=#ffffff align=center>
<table width="100%" border="0" cellspacing="0" cellpadding="1" bgcolor="#99CCCC">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
        <tr>
          <td>
            <table width="100%" border="0" cellspacing="3" cellpadding="3" height="340" align=center>
              <tr bgcolor="#006699"> 
                <td align="center" colspan="2" height="30" bgcolor="#006699"> 
                <%
				 response.write "<font color=#ffffff><b>"&server.htmlencode(title)&"</b></font>"
                %>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">
                  <table width="100%" border="0" cellspacing="2" cellpadding="0">
                    <tr> 
                      <td height="23">&nbsp;&nbsp;&nbsp;&nbsp;
                      <%
                        response.write replace(server.htmlencode(content),vbcrlf,"<br>")
                      %>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor="#006699" align="center">
                <td colspan="2" bgcolor="#006699" height="30" align="right">
                  <%response.write "<font color=#ffffff>Date : "&(createtime)&"</font>&nbsp;"%>
                </td>
              </tr>
              <tr> 
                <td align="center" height="10" colspan="2"><A href="javascript:window.close();">Close 
                  Window</A></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<div align="center"><script language=JavaScript src="included/copyright.js"></script></div>
</BODY></HTML>