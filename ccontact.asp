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
		sqlcmd=sqlcmd&"CompanyName,comment,create_date) "
		sqlcmd=sqlcmd&"values ('"& title &"','"& fax 
		sqlcmd=sqlcmd&"','"& phone &"','"& fn
		sqlcmd=sqlcmd &"','"& address &"','"& em 
		sqlcmd=sqlcmd&"','"& cn &"','"& comment &"',#"& dat &"#)"
		'response.write sqlcmd
	    conn.Execute sqlcmd
	end if 
	sqlcmd="select*from member where email='"&em&"'"
	set rs=conn.execute(sqlcmd)
	'response.write rs.eof&"ppppp"
   if  not rs.eof then 
     session("user_id")=rs("id")
     response.redirect "sendemail.asp?email="&server.Urlencode(email)
	end if  
end if

%>

<!DOCTYPE html>
<html lang="en">
<HEAD>
<meta charset="Big5"> 
<meta name="keywords" content="GC-heat, EOC , Hasco, strack, DME, progressive, Temperature Controller, band heaters, coil heaters, date inserts, Thermocouple, sensor, shot counter, hydraulic cylinder, mold and die components and functional elements, cooling system, industrial heating, vega;wema;hotset;opitz;GC-Heater;加熱圈;發熱圈;發熱棒;溫度控制器;熱電偶;加熱器;加熱管;發熱筒;油缸;溫控箱;享基;享基科技有限公司;">
<TITLE>Enhance Techologies Limited</TITLE>
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">

<!-- jQuery library -->
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script>

<!-- Latest compiled JavaScript -->
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>

<!-- responsive to mobile devices -->
<meta name="viewport" content="width=device-width, initial-scale=1">

<!--#include file="menu_bar.js"-->

function dosubmit()
{
	var errmsg;
	var strmail;
	errmsg="";
	strmail="";
	 strpwd=document.updata.Product.value;
	if(strpwd=="")
	errmsg=errmsg+"Please Input Product& Qty.\n";	
	if(document.updata.name.value=="")
	errmsg=errmsg+"Please Input  name.\n";
	if(document.updata.title.value=="")
	errmsg=errmsg+"Please Input  title.\n";	
    if(document.updata.company.value=="")
	errmsg=errmsg+"Please Input company Name.\n";
	if(document.updata.address.value=="" )
	errmsg=errmsg+"Please Input address.\n";
	strmail=document.updata.email.value;
	if(strmail=="" )
		errmsg=errmsg+"Please Input e-mail address.\n";
	else
		if(strmail.search("@")==-1)
			errmsg=errmsg+"Please Input a right e-mail address.\n";
	//phone=document.updata.phone.value
	//if(phone=="" )
	//	errmsg=errmsg+"Please Input phone.\n";
	//else
	//	if(isNaN(phone))
	//		errmsg=errmsg+"Please Input a right phone.\n";
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
<BODY>

<div class="container">


<!--#include file="menu_nav.inc"-->
 
<table border="0" cellspacing="0" cellpadding="0" width="100%" height="455">
  <tr align="center"> 
    <td colspan="2" height="125"> 
      <!--#include file="top.inc" -->
    </td>
  </tr>
  <tr> 
    <td valign="top" width="145" align="center" rowspan="3"> 
      <!--#include file="cleft.inc" -->
    </td>
    <td  width="644" valign="top"> 
      <form name="updata" mothed="post" action="ccontact.asp">
        <div align="center">
        <div align="center"><center><p><font face="Arial" size="2"
          color="#EFFBBF"><strong>如閣下想知多一些本公司的資料,
          請在下列輸入閣下的資料.<br>
			我們將寄上產品光碟.</strong></font></p>
          </center></div><div align="center"><center><table border="1" width="588">
            <tr>
              <td align="right" width="128" valign="top"><strong><font face="Arial" size="2"
              color="#EFFBBF">產品 &amp; 數量 :<br>
              </font><font face="Arial" color="#EFFBBF" size="1">e.g. 型號.,<br>
              Metal Ball Pen</font></strong></td>
			  
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
              color="#EFFBBF" face="Arial"><strong>(必需填寫)</strong></font></small></small></td>
<% else %>
                  <td width="448"><small>
                    <textarea rows="5" name="Product" cols="36"></textarea>
                    <small><font
              color="#EFFBBF" face="Arial"><strong>(必需填寫)</strong></font></small></small></td>
<% end if%>			  		
            </tr>
            <tr>
              <td width="128" align="right" valign="top">&nbsp;</td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF"><strong><input type="checkbox"
              name="Send_Product_Literature" value="ON"> 要求產品印刷品</strong></font></td>
            </tr>
            <tr>
              <td width="128" align="right" valign="top">&nbsp;</td>
              <td valign="top" width="448"><font face="Arial" size="2" color="#EFFBBF"><strong><input
              type="checkbox" name="Send_CD" value="ON"> 要求光碟</strong></font></td>
            </tr>
            <tr>
              <td width="128" align="right" valign="top">&nbsp;</td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF"><strong><input type="checkbox"
              name="Have_a_Salesperson_Contact_Me" value="ON">請聯絡我</strong></font></td>
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
              <td align="right" width="128" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>名稱
              :</strong></font></td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF">
                    <input type="text" size="32"
              name="name" value="<%=name%>">
                    </font><font face="Arial" color="#EFFBBF" size="1"><strong>(必需填寫)</strong></font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>職銜
              :</strong></font></td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF">
                    <input type="text" size="32"
              name="title" value="<%=title%>">
                    </font><font face="Arial" color="#EFFBBF" size="1"><strong>(必需填寫)</strong></font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>公司
              :</strong></font></td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF">
                    <input type="text" size="32"
              name="company" value="<%=company%>">
                    </font><font face="Arial" color="#EFFBBF" size="1"><strong>(必需填寫)</strong></font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>地址
              :</strong></font></td>
                  <td width="448"><font face="Arial" size="2" color="#EFFBBF">
                    <textarea name="address"
              rows="5" cols="31"><%=address%></textarea>
                    </font><font face="Arial" color="#EFFBBF" size="1"><strong>(必需填寫)</strong></font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>電郵
              :</strong></font></td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF">
                    <input type="text" size="25"
              name="email" value="<%=email%>">
                    </font><font face="Arial" color="#EFFBBF" size="1"><b>(必需填寫)</b></font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>電話
              :</strong></font></td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF">
                    <input type="text" size="25"
              name="phone" value="<%=phone%>">
                    </font><font face="Arial" color="#EFFBBF" size="1"><strong>(必需填寫)</strong></font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>傳真
              :</strong></font></td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF">
                    <input type="text" size="25"
              name="fax" value="<%=fax%>">
                    </font><font face="Arial" color="#EFFBBF" size="1"><strong>(必需填寫)</strong></font></td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top"><strong><font face="Arial" size="2"
              color="#EFFBBF">意見  &amp; 信息:</font></strong></td>
              <td width="448"><font face="Arial" size="2" color="#EFFBBF"><strong><textarea rows="5"
              name="Comment" cols="31"></textarea></strong></font></td>
            </tr>
            <tr>
              <td width="128" align="right" valign="top">&nbsp;</td>
              <td valign="bottom" width="448"><font face="Arial" size="2" color="#EFFBBF">
			  <input type="button" value="提交要求" onClick="javascript:dosubmit();"> 
			  <input type="reset" value="重新填寫"> 
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
        <td width="250"></td>
        <td width="343"></td>
    </tr>
    <tr>
        <td width="250"><font face="Arial" size="2" color="#C0C0C0"><strong><u>Head 
          Office</u> </strong></font></td>
        <td width="343"></td>
    </tr>
    <tr>
        <td width="250"><font face="Arial" size="2" color="#C0C0C0"><strong> G/F, 
          On Tai Bldg.,<br>
        47-49 Ting On Street, <br>
        Nagu Tau Kok, H.K.<br>
    Tel: (852) 2798 6263 (12 lines)<br>
    Fax: (852) 2751 6659 (2 lines)
    </strong></font></td>
        <td width="343" valign="top"><font face="Arial" size="2" color="#C0C0C0"><b>Flat 
          C, 1/F., Kam Mong Bldg.<br>
        37-39 Fa Yuen street, Kln. H.K.<br>
    Tel: (852) 2771 1333, 2780 1029<br>
    Fax: (852) 2770 0351</b></font></td>
    </tr>
    <tr>
        <td colspan="2"> 
          <p align="center"><font face="Arial" size="2" color="#C0C0C0"><strong>
    E-mail: <a href="mailto:info@jaguarpen.com">info@jaguarpen.com</a>
    </strong></font></td>
    </tr>
  </table>
  </center>
</div>

</div>


<!--#include file="menu_bottom.inc"-->

</div>
</BODY>
</html>
