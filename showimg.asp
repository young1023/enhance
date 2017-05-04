<!--#include file="database.inc"-->
<%
product_id=trim(request("id"))
flage=trim(request("flage"))

if flage=1 then
  mysql="select picpath from product where id="& product_id &" "
  Set v= Server.CreateObject("ADODB.Recordset")
  v.Open mysql,conn,1,1

  if not v.eof then
    if len(v("picpath")) then
      Response.ContentType = "image/gif"
      Response.BinaryWrite v("picpath").getChunk(7500000)
    else
      response.redirect "images/SSKEY7.gif"
    end if

    v.close
    set v=nothing

  end if

elseif flage=2 then
  mysql="select pic2path from product where id="& product_id &" "
  Set v= Server.CreateObject("ADODB.Recordset")
  v.Open mysql,conn,1,1

  if not v.eof then
    if len(v("pic2path")) then
      Response.ContentType = "image/gif"
      Response.BinaryWrite v("pic2path").getChunk(7500000)
    else
      response.redirect "images/SSKEY7.gif"
    end if

    v.close
    set v=nothing

  end if
end if
%>