<!--#include file="database.inc"-->

<%
act_bt=request("txtmail")
straction=request("action_button")
email=request("email")

if straction="save" then         
         fn=replace(request("name"),"'","''")
         title=replace(request("title") ,"'","''")
         em=replace(request("email"),"'","''")      
		 cn=replace(request("company"),"'","''")
         address=replace(request("address"),"'","''")
		 fax=replace(request("fax"),"'","''")
		 phone=replace(request("phone"),"'","''")
         moble=replace(request("moble"),"'","''")
		 product=replace(request("Product"),"'","''")
		 session("product")=product
        comment=replace(request("comment"),"'","''")				 
		if trim((comment)&"")="" then
		    comment=null
		end if
		dat=date()
	sqlcmd="select*from member where email='"&em&"'"
    'response.write sqlcmd
	set rs=conn.execute(sqlcmd) 
	
	if not rs.eof then
        strsql1="update member set firstname= '"& fn & "' ," _
			          &"title='"& title &"'," _
			          &"phone='"& phone &"'," _
			          &"fax='"& fax & "'," _
			          &"CompanyName='"& cn & "', " _
					  &"comment='"& comment & "' ," _
					   &"create_date=#"& dat & "# ," _
					   &"address='"& address & "' " _
					   &"WHere email='"& em&"'  "
	     	'response.write strsql1
			     conn.execute strsql1
			     set strsql1=nothing 
   end if
   if rs.eof then 
		sQLcmd= "insert into member( title,"
		sqlcmd=sqlcmd&"fax,phone,Firstname,"
		sqlcmd=sqlcmd&"address,Email,"
    	sqlcmd=sqlcmd&"moble,"
		sqlcmd=sqlcmd&"CompanyName,comment,create_date) "
		sqlcmd=sqlcmd&"values ('"& title &"','"& fax 
		sqlcmd=sqlcmd&"','"& phone &"','"& fn
		sqlcmd=sqlcmd &"','"& address &"','"& em & "','"& moble 
		sqlcmd=sqlcmd&"','"& cn &"','"& comment &"',#"& dat &"#)"
		'response.write sqlcmd
	    conn.Execute sqlcmd
	end if 
	sqlcmd="select*from member where email='"&em&"'"
	set rs=conn.execute(sqlcmd)
	'response.write rs.eof&"ppppp"
   if  not rs.eof then 
     session("user_id")=rs("id")
    ' response.redirect "sendemail.asp?email="&server.Urlencode(email)
 response.redirect "ssearchresult.asp"

	end if  
end if
if straction=""then
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GB2312">
<meta name="description"
content="manufacturer of pens, writing instruments, stationery, gifts, ball pens, premiums &amp; incentives, jaguar pens giftware &amp; novelties,  pens &amp; pen sets">
<meta name="keywords"
content="manufacturer, pens, writing instruments, WDC, stationery, gifts, ball pens, premiums &amp; incentives, jaguar pens giftware &amp; novelties, Winson, pens &amp; pen sets">
<title>Contact Us</title>
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
function dosubmit()
{
	var errmsg;
	var strmail;
	errmsg="";
	strmail="";
	 strpwd=document.updata.Product.value;
	if(strpwd=="")
	errmsg=errmsg+"请输入产品及数量.\n";	
	if(document.updata.name.value=="")
	errmsg=errmsg+"请输入名称.\n";
	if(document.updata.title.value=="")
	errmsg=errmsg+"请输入职衔.\n";	
    if(document.updata.company.value=="")
	errmsg=errmsg+"请输入公司名称.\n";
	if(document.updata.address.value=="" )
	errmsg=errmsg+"请输入地址.\n";
	strmail=document.updata.email.value;
	if(strmail=="" )
		errmsg=errmsg+"请输入电邮.\n";
	else
		if(strmail.search("@")==-1)
			errmsg=errmsg+"请输入正确的电邮.\n";
	//phone=document.updata.phone.value
	//if(phone=="" )
	//	errmsg=errmsg+"请输入电话.\n";
	//else
	//	if(isNaN(phone))
	//		errmsg=errmsg+"请输入正确的电话.\n";
	//fax=document.updata.fax.value
	//if(fax=="" )
	//	errmsg=errmsg+"Please Input FAX.\n";
	//else
	//	if(isNaN(fax))
	//		errmsg=errmsg+"Please Input a right FAX.\n";  
	
	if(errmsg!="")
	{
		alert(errmsg);
	}
	else
	{
		document.updata.action_button.value="save";
		document.updata.submit();
	}
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="images/demo/bdg1.jpg" link="ffffff" vlink="ffffcc">
 
<table border="0" cellspacing="0" cellpadding="0" width="100%" height="455">
  <tr align="center"> 
    <td colspan="2" height="125"> 
   <p align="center"><img src="images/layout-header_c.jpg"></p>
    </td>
  </tr>
  <tr> 
    <td valign="top" width="145" align="center" rowspan="3"> 
      <!--#include file="sleft.inc" -->
    </td>
    <td  width="644" valign="top"> 
      <form name="updata" mothed="post" action="scontact.asp">
        <div align="center">
        <div align="center"><center><p><font face="Arial" size="2"
          color="#EFFBBF">如阁下想知多一些本公司的資料, 请在下列输入阁下的資料.<br>我们将寄上产品光盘</font></p>
          </center></div><div align="center"><center><table border="1" width="588">
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" size="2"
              color="#EFFBBF">产品 & 数量:<br>
              </font><font face="Arial" color="#EFFBBF" size="1">e.g.型号<br>
              Metal Ball Pen</font></td>
			  
 <%
  if session("enquiryno")<>"" then
 	strsql="Select * from Acclist where A_id= "& session("enquiryno") &" ORDER BY Acclist.id DESC  "	
	
	set acres=conn.execute(strsql)				
	if not Acres.eof then 
	Do while (not acres.eof)      
         strtemp=strtemp&chr(10)&trim(Acres("P_model")&"")  &space(3)& Acres("qty")   
         acres.movenext					
    loop  
    end if
%>
      <td width="448"><small><textarea rows="5" name="Product" cols="36"><%=strtemp%></textarea><small><font
              color="#EFFBBF" face="Arial">(必需填写)</font></small></small></td>
<% else %>
                  <td width="448"><small>
                    <textarea rows="5" name="Product" cols="36"></textarea>
                    <small><font
              color="#EFFBBF" face="Arial">(必需填写)</font></small></small></td>
<% end if%>			  		
            </tr>
            <tr>
              <td width="128" align="right" valign="top">&nbsp;</td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF"><input type="checkbox"
              name="Send_Product_Literature" value="ON"> 要求产品印刷品</font></td>
            </tr>
            <tr>
              <td width="128" align="right" valign="top">&nbsp;</td>
              <td valign="top" width="448"><font face="Arial" size="2" color="#EFFBBF"><input
              type="checkbox" name="Send_CD" value="ON">要求光盘</font></td>
            </tr>
            <tr>
              <td width="128" align="right" valign="top">&nbsp;</td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF"><input type="checkbox"
              name="Have_a_Salesperson_Contact_Me" value="ON">请联络我</font></td>
            </tr>
			<%
	 if session("user_id")<>"" then
			 sql="select lastname,email,title,companyname,address,phone,fax from member where id="& session("user_id") &" "
             set w=conn.execute(sql)
             if not w.eof then 
			   name=w("lastname")
			   title=w("title")
			   company=w("companyname")
			   address=w("address")
			   email=w("email")
			   phone=w("phone")
			   fax=w("fax")
			  end if
	  end if
			%>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" size="2" color="#EFFBBF">名称
              :</font></td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF">
                    <input type="text" size="32"
              name="name" value="<%=name%>">
                    </font><font face="Arial" color="#EFFBBF" size="1">(必需填写)</font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" size="2" color="#EFFBBF">职衔 :</font></td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF">
                    <input type="text" size="32"
              name="title" value="<%=title%>">
                    </font><font face="Arial" color="#EFFBBF" size="1">(必需填写)</font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" size="2" color="#EFFBBF">公司
              :</font></td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF">
                    <input type="text" size="32"
              name="company" value="<%=company%>">
                    </font><font face="Arial" color="#EFFBBF" size="1">(必需填写)</font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" size="2" color="#EFFBBF">地址  :</font></td>
                  <td width="448"><font face="Arial" size="2" color="#EFFBBF">
                    <textarea name="address"
              rows="5" cols="31"><%=address%></textarea>
                    </font><font face="Arial" color="#EFFBBF" size="1">(必需填写)</font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" size="2" color="#EFFBBF">电邮
              :</font></td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF">
                    <input type="text" size="25"
              name="email" value="<%=email%>">
                    </font><font face="Arial" color="#EFFBBF" size="1">(必需填写)</font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" size="2" color="#EFFBBF">公司电话
              :</font></td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF">
                    <input type="text" size="25"
              name="phone" value="<%=phone%>">
                    </font><font face="Arial" color="#EFFBBF" size="1">(必需填写)</font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" size="2" color="#EFFBBF">公司传真
              :</font></td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF">
                    <input type="text" size="25"
              name="fax" value="<%=fax%>">
                    </font><font face="Arial" color="#EFFBBF" size="1">(必需填写)</font></td>
            </tr>
           <tr>
              <td align="right" width="128" valign="top"><font face="Arial" size="2" color="#EFFBBF">手机
              :</font></td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF">
                    <input type="text" size="25"
              name="fax" value="<%=moble%>">
                    </font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" size="2"
              color="#EFFBBF">意见 & 信息:</font></td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF"><strong><textarea rows="5"
              name="Comment" cols="31"></textarea></strong></font></td>
            </tr>
            <tr>
              <td width="128" align="right" valign="top">&nbsp;</td>
              <td valign="bottom" width="448"><font face="Arial" size="2" color="#EFFBBF">
			  <input type="button" value="提交" onClick="javascript:dosubmit();"> 
			  <input type="reset" value="重新填写"> 
			  <input type="hidden" name="action_button" value=""></font></td>
            </tr>
          </table></center></div></div>
		  </form>
    </td>
  </tr>
  <tr>
    <td  width="644" valign="top">&nbsp;</td>
  </tr>
</table>
<div align="center">
  <center>
    <table border="0" cellpadding="0" cellspacing="0" width="593">
      <tr>
        <td width="350"></td>
        <td width="343"></td>
       <td width="343"></td>
    </tr>
    <tr>
        <td width="350"><font face="Arial" size="2" color="#C0C0C0"><strong><u>总公司</u> </strong></font></td>
        <td width="343"><font face="Arial" size="2" color="#C0C0C0"><strong><u>旺角</u> </strong></font></td>
<td width="343"><font face="Arial" size="2" color="#C0C0C0"><strong><u>深圳</u> </strong></font></td>
    </tr>
    <tr>
        <td width="350"><font face="Arial" size="2" color="#C0C0C0">香港九龙<br>
观塘道330号<br>威力货运大厦1楼。
         <br>
    电话: (852) 2798 6263 (12 lines)<br>
    传真: (852) 2751 6659 (2 lines)
   </font></td>
        <td width="343"><font face="Arial" size="2" color="#C0C0C0">香港九龙<br>
花园街37-39号<br>金旺楼一楼C座。<br>
    电话: (852) 2771 1333, 2780 1029<br>
    传真: (852) 2770 0351</font></td>
<td width="343"><font face="Arial" size="2" color="#C0C0C0">深圳市罗湖区
文锦中路1043号<br>
联兴大厦南座812室<br>
电话：0755-8225 1929<br>
传真：0755-8225 8250
</font></td>
    </tr>
    <tr>
        <td colspan="3"> 
          <p align="center"><font face="Arial" size="2" color="#C0C0C0"><strong>
    电邮: <a href="mailto:info@jaguarpen.com">info@jaguarpen.com</a>
    </strong></font></td>
    </tr>
  </table>
  </center>
</div>

<p align="center"><font color="#C0C0C0"><font face="Arial"> <br>
  <small><small>Copyright</small></small></font><font face="Times New Roman"><font size="2"> 
  &copy; 2001</font></font><font
face="Arial" color="#C0C0C0"><small><small>&nbsp;&nbsp;&nbsp; </small></small></font><font
face="Arial" size="2" color="#C0C0C0">Winson Development Company</font><font face="Arial"
color="#C0C0C0"><small><small>&nbsp;&nbsp;&nbsp;&nbsp; All right reserved.</small></small></font></font></p>

</body>
</html>
<% end if%>