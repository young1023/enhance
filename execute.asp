<%
if session(("SECURITY_LEVEL"))<>"1" then
  response.redirect "default.asp"
end if
session("picflage")=""
session("product_id")=""
%>
<!--#include file="included/conne.inc"-->
<%
response.expires=0
flage=trim(request.form("whatdo"))
'response.write flage
'response.end


if flage="muinfo" then   'modify user information
  name=replace(trim(request.form("name")),"'","''")
  indicate=replace(trim(request.form("indicator")),"'","''")   'region
  phone=replace(trim(request.form("phone")),"'","''")          'department
  sql="update member set name='"&name&"',indicate='"&indicate&"',phone='"&phone&"' where id="&session("shell_id")&" "
  conn.execute(sql)
  whatgo="user_info.asp"
  message="Update Successfully"
elseif flage="userpost" then
  modulnum=trim(request.form("modulnum"))
  if modulnum="" then
    conn.close
    set conn=nothing
    response.redirect "listbill.asp"
  end if
  terminal=replace(trim(request.form("terminal")),"'","''")
  if request.form("u_select")=1 then
    country=trim(request.form("u_region"))
    sql="select mnemonic from country where location='"&country&"'"
    set rs2=conn.execute(sql)
    if rs2.eof then
      sql="Insert into country(country,location,mnemonic) values('"&country&"','"&country&"','"&country&"')"
      conn.execute(sql)
    else
      country=rs2("mnemonic")
    end if
    rs2.close
    set rs2=nothing
  else
    country=trim(request.form("country"))
  end if
  money=split(trim(request.form("money")),",")
  money2=split(trim(request.form("money2")),",")
  for i=0 to Ubound(money)
    if trim(money(i))<>"" then
      p1=trim(money2(i))+","+p1
      p2=trim(money(i))+","+p2
    end if
  next
 if p1<>"" and p2<>"" and terminal<>"" then
  p1=p1&"userid,terminal,flage"
  p2=p2&session("shell_id")&",'"&terminal&"','"&country&"'"
  sql="Insert into "&modulnum&"_b("&p1&") values("&p2&")"
  'response.write sql
  'response.end
  conn.execute(sql)
  message="Post Value Successfully"
 else
  message="Post Value Error"
 end if
  whatgo="userpost.asp?id="&modulnum
  whatgo="window.close()"
elseif flage="addmember" then
  pid=trim(request.form("pid"))
  name=replace(trim(request.form("name")),"'","''")
  employeenum=replace(trim(request.form("employeenum")),"'","''")
  indicator=replace(trim(request.form("indicator")),"'","''")
  if indicator="" then
    indicator=" "
  end if
  phone=replace(trim(request.form("phone")),"'","''")
  if phone="" then
    phone=" "
  end if

  sql="select id from member where employeenum='"&employeenum&"' "
  set rs=conn.execute(sql)
  if not rs.eof then
     message="The Employee Number <u><b>"&employeenum&"</b></u> is exist !"
  else
     sql="insert into member(name,pwd,employeenum,indicate,phone,flage) values('"&name&"','"&employeenum&"','"&employeenum&"','"&indicator&"','"&phone&"',"&pid&")"
     conn.execute(sql)
     message="Add A Member Successfully"
  end if
  rs.close
  set rs=nothing
  whatgo="sa_member.asp"
elseif flage="modifymember" then
  pid=trim(request.form("pid"))
  id=trim(request.form("id"))
  name=replace(trim(request.form("name")),"'","''")
  employeenum=replace(trim(request.form("employeenum")),"'","''")    'e-mail
  employeenum2=replace(trim(request.form("employeenum2")),"'","''")
  indicator=replace(trim(request.form("indicator")),"'","''")
  if indicator="" then
    indicator=" "
  end if
  phone=replace(trim(request.form("phone")),"'","''")
  if phone="" then
    phone=" "
  end if

  sql="select id from member where employeenum='"&employeenum&"' "
  set rs=conn.execute(sql)
  if not rs.eof then
     if employeenum=employeenum2 then
       sql="Update member set name='"&name&"',indicate='"&indicator&"',phone='"&phone&"',flage="&pid&" where id="&id&" "
       conn.execute(sql)
       message="Modify Member Successfully"
     else
       message="The Employee Number <u><b>"&employeenum&"</b></u> is exist !"
     end if
  else
     sql="Update member set name='"&name&"',pwd='"&employeenum&"',employeenum='"&employeenum&"',indicate='"&indicator&"',phone='"&phone&"',flage="&pid&" where id="&id&" "
     conn.execute(sql)
     message="Modify Member Successfully"
  end if
  rs.close
  set rs=nothing
  whatgo="sa_modify.asp?id="+id+"&yn=1"

'==================
'
'Add News
'
'==================

elseif flage="addnews" then
  title=replace(trim(request.form("title")),"'","''")
  content=replace(trim(request.form("content")),"'","''")
  postdate=replace(trim(request.form("postdate")),"'","''")
  news_id=replace(trim(request.form("news_id")),"'","''")
  news_link=replace(trim(request.form("news_link")),"'","''")
'response.write postdate
'response.write news_id
'response.end

 
  sql="insert into message(title,http, createtime,typeid) values('"&title&"','"&news_link&"','"&postdate&"','"&news_id&"')"
  conn.execute(sql)
  whatgo="sa_news.asp"
  message="News was added successfully"

'==================
'
'Add Show
'
'==================

elseif flage="addshow" then
  trade_name=replace(trim(request.form("trade_name")),"'","''")
  trade_link=replace(trim(request.form("trade_link")),"'","''")
  trade_date=replace(trim(request.form("trade_date")),"'","''")
  trade_id=replace(trim(request.form("trade_id")),"'","''")
  trade_location=replace(trim(request.form("trade_location")),"'","''")
'response.write postdate
'response.write news_id
'response.end

 sql="insert into tradeshow(trade_name,location,trade_date,http,trade_id) values('"&trade_name&"','"&trade_location&"','"&trade_date&"','"&trade_link&"','"&trade_id&"')"
  conn.execute(sql)
  whatgo="sa_tradeshow.asp"
  message="Trade Show was added successfully"

'===================
'
' Edit Newsletter
'
'===================



elseif flage="mnews" then
  id=trim(request.form("id"))
  title=replace(trim(request.form("title")),"'","''")
  news_link=replace(trim(request.form("news_link")),"'","''")
  'content=replace(trim(request.form("content")),"'","''")
  Typeid=replace(trim(request.form("typeid")),"'","''")
  postdate=replace(trim(request.form("postdate")),"'","''")
'response.write postdate
'response.write typeid
'response.end


  sql="Update message set title='"&title&"',http='"&news_link&"',createtime='"&postdate&"',typeid='"&typeid&"' where id="&id&" "
  conn.execute(sql)
  whatgo="sa_news.asp?id="+id+"&yn=1"
  message="Modify News Successfully"


'===================
'
' Edit Tradeshow
'
'===================



elseif flage="mtradeshow" then
  id=trim(request.form("id"))
  title=replace(trim(request.form("title")),"'","''")
  location=replace(trim(request.form("location")),"'","''")
  Trade_id=replace(trim(request.form("typeid")),"'","''")
  Trade_link=replace(trim(request.form("trade_link")),"'","''")
  trade_date=replace(trim(request.form("postdate")),"'","''")
'response.write postdate
'response.write trade_link
'response.end


  sql="Update tradeshow set trade_name='"&title&"',location='"&location&"',http='"&trade_link&"',trade_date='"&trade_date&"',trade_id='"&trade_id&"' where id="&id&" "
  conn.execute(sql)
  whatgo="sa_tradeshow.asp?id="+id+"&yn=1"
  message="Modify Tradeshow Successfully"




elseif flage="changepwd" then
  oldpwd=trim(request.form("oldpwd"))
  newpwd=trim(request.form("newpwd"))
  sql="select id from member where pwd='"&oldpwd&"' and id="&session("shell_id")&" "
  set rs=conn.execute(sql)
  if not rs.eof then
    sql="update member set pwd='"&newpwd&"' where id="&session("shell_id")&" "
    conn.execute(sql)
    message="<font color=white>Change Password Successfully</font>"
    whatgo="user_info.asp"
  else
    message="Login Password Is Wrong"
    whatgo="user_pwd.asp"
  end if
  rs.close
  set rs=nothing
else
  response.redirect "index.asp"
end if


conn.close
set conn=nothing
wait=3

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Refresh" content="<%=wait%>; url='<%=whatgo%>'">
<title>Execute Page</title>
</head>
<body topmargin="0" marginwidth="0" marginheight="0" leftmargin="0" >
<br><br>
<table border=0 cellpadding=3 cellspacing=0 class=hardcolor width="90%" height="100%" align=center>
  <tbody>
  <tr> 
    <td align=center bgcolor="#006699" height="28"><font color=white><%=message%></font></td>
  </tr>
  <tr>
    <td align=center height="38"><br><a href='<%=whatgo%>'>Return</a></td>
  </tr>
  </tbody>
</table>
</body>
</html>