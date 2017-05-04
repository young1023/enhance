<!--#include file="database.inc"-->
<%
if session("user_id")="" then
	response.redirect "login.asp"
end if
struid=session("user_id")
straction=request("action_button")
ps=trim(request("txtpass"))
rps=trim(request("txtrepass"))
qt=trim(request("selquestion"))
aw=trim(request("txtanswer"))
fn=trim(request("txtfirstname"))
ln=trim(request("txtlastname"))
em=trim(request("txtemail"))
gd=trim(request("sex"))
sb=trim(request("sub"))
ct=trim(request("selstate"))
cn=trim(request("txtcompanyname"))
cs=trim(request("selsize"))
bt=trim(request("seltype"))
sd=trim(request("txtsaddress"))
if straction="save"then
	strerr=""
	sqlcmd="select * from member where email='"&em&"' and id<>"& struid & " "
	'response.write sqlcmd
	set rs=conn.execute(sqlcmd) 
	if not rs.eof then
		strerr="Sorry,This email is exist!<br>You should input another email."
	else
		ps=replace(request("txtpass"),"'","''")
		qt=replace(request("selquestion"),"'","''")
		aw=replace(request("txtanswer") ,"'","''")
		fn=replace(request("txtfirstname"),"'","''")
		ln=replace(request("txtlastname") ,"'","''")
		em=replace(request("txtemail"),"'","''")
		gd=request("sex")
		sb=request("sub")
		ct=request("selstate")
		cn=replace(request("txtcompanyname"),"'","''")
		cs=replace(request("selsize"),"'","''")
		bt=request("seltype")
		sd=replace(request("txtsaddress"),"'","''")
	 'straction=trim(request("Action_Button"))
	 sqlcmd="update member set mypassword='"&ps&"',question='"&qt&"',answer='"&aw&"',firstname='"&fn&"',lastname='"&ln&"',sex='"&gd&"',Country='"&ct&"',CompanyName='"&cn&"',ComanySize='"&cs&"',BusinessType='"&bt&"',subscribe="&sb&",email='"&em&"' where id="& struid &" "
	 response.write "<br>"&sqlcmd
	   conn.Execute(sqlcmd)
		if err.number>0 then
			strerr="Save memeber's information raise error! "
		else
			sqlcmd="select*from member where email='"&em&"'"
			set rs=conn.execute(sqlcmd)
			if  not rs.eof then 
				session("id")=rs("id")
				session("level")=rs("level")
			end if  
			strerr="Update memeber's information success! "
		end if
	end if
end if
if struid<>"" then
	strsql="select * from member where id="& struid
	set acres=conn.execute(strsql)
	if not acres.eof then
		ps=trim(acres("mypassword")&"")
		rps=trim(acres("mypassword")&"")
		qt=trim(acres("question")&"")
		aw=trim(acres("answer")&"")
		fn=trim(acres("firstname")&"")
		ln=trim(acres("lastname")&"")
		em=trim(acres("email")&"")
		gd=trim(acres("sex")&"")
		sb=trim(acres("subscribe")&"")
		ct=trim(acres("country")&"")
		cn=trim(acres("CompanyName")&"")
		cs=trim(acres("comanysize")&"")
		bt=trim(acres("businesstype")&"")
		sd=trim(acres("saddress"))
	end if
end if
function fillcob(strtype,strid)
	dim acres,strsql,acid,strsel
	select case strtype
		case "question"
			strsql="select id,question as itemval from question order by id"
		case "country"
			strsql="select id,country as itemval from country order by country"
		case "csize"
			strsql="select id,companysize as itemval from companysize order by id"
		case "btype"
			strsql="select id,businesstype as itemval from businesstype order by id"
	end select
	set acres=conn.execute(strsql)
	do while not acres.eof
		 acid=trim(acres("itemval")&"")
		 strsel=""
		 if strid<>"" then
			 if acid=strid then
			 	strsel="selected"
			 end if
		 end if
		 response.write "<option value="&chr(34)&acid&chr(34)&" "&strsel&">"&trim(acres("itemval")&"")&"</option>"
		 acres.movenext
	loop
	set acres=nothing
end function
%>
<html>
<head>
<title>Add Member</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="images/Tabcss.css" type="text/css">
<script language="JavaScript">
<!--
function dosubmit()
{
	var errmsg;
	var strmail;
	var strpwd;
    var strrpwd;
	errmsg="";
	strmail="";
	strmail=document.updata.txtemail.value;
	if(strmail=="" )
		errmsg=errmsg+"Please Input e-mail address.\n";

	else
		if(strmail.search("@")==-1)
			errmsg=errmsg+"Please Input a right e-mail address.\n";
	strpwd=document.updata.txtpass.value;
	strrpwd=document.updata.txtrepass.value;
	if(strpwd=="")
		errmsg=errmsg+"Please Input Password.\n";
 	else{
		if(strpwd.length<5 )
			errmsg=errmsg+"Password is short!Password length is greater than 5!\n";
		else
			{
			if(strrpwd=="")
				errmsg=errmsg+"Please Input Confirm Password.\n";
			else	 
				if(strrpwd.length<5 )
					errmsg=errmsg+"Confirm Password is short!Password length is greater than 5!\n";
				else
			 		if(strpwd!=strrpwd)
						errmsg=errmsg+"Password is not equal Confirm Password!\n";
					}		
			}
	if(document.updata.selquestion.value=="")
		errmsg=errmsg+"Plesase Select Question!\n";
	if(document.updata.txtanswer.value=="")
		errmsg=errmsg+"Plesase Input Answer!\n";
	if(document.updata.sex.value=="")
		errmsg=errmsg+"Please Input gender.\n";
	if(document.updata.selstate.value=="")
		errmsg=errmsg+"Plesase Select Country!\n";
	if(document.updata.selsize.value=="")
		errmsg=errmsg+"Plesase Select Company Size!\n";
	if(document.updata.txtcompanyname.value=="")
		errmsg=errmsg+"Plesase Input Company Name!\n";
	if(document.updata.seltype.value=="")
		errmsg=errmsg+"Plesase Select Company Type!\n";

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
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="images/demo/bdg1.jpg">
<table border="0" cellspacing="0" cellpadding="0" width="100%" height="455">
  <tr align="center"> 
    <td colspan="2" height="125"> 
      <!--#include file="top.inc" --></td>
  </tr>
  <tr> 
    <td valign="top" width="145" align="center" rowspan="2"> 
      <!--#include file="left.inc" --></td>
    <td height="103" width="644" valign="top"> 
      <!--#include file="sbar.inc" --></td>
  </tr>
  <tr>
    <td width="644" valign="top">
      <form  name="updata" method="post" action="myprofile.asp">
        <table
      border="0" width="694">
          <tr> 
            <td align="center" width="646"> 
              <div align="center">
                <center>
                  <p><font face="Arial" size="2"
          color="#EFFBBF"><strong><br>
                    Please fill out the form below !<br>
                    <font color="Red"><%=strerr%></font> </strong></font></p>
                </center>
              </div>
        </table>
        <table border="1" width="472"align="center">
          <tr> 
            <td align="right" width="149" valign="top" height="20"><strong><font face="Arial" size="2" color="#EFFBBF">E_mail:</font></strong></td>
            <td width="307" height="20"><small> 
              <input type="text" name="txtemail" maxlength="50"size="24"  value="<%=em%>">
              <font face="Arial" color="#EFFBBF" size="1"><strong>(required)</strong></font></small> 
            </td>
          </tr>
          <tr> 
            <td align="right" width="149" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>Password:</strong></font></td>
            <td width="307"><font face="Arial" size="2" color="#EFFBBF"> </font><font face="Arial" color="#EFFBBF" size="1"><strong><small> 
              <input type="password" name="txtpass" maxlength="50"size="24"  value="<%=ps%>">
              </small>(required) </strong></font></td>
          </tr>
          <tr> 
            <td align="right" width="149" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>Re_Enter 
              password:</strong></font></td>
            <td width="307"><font face="Arial" size="2" color="#EFFBBF"> <small> 
              <input type="password" name="txtrepass" maxlength="50"size="24"  value="<%=rps%>">
              </small> </font><font face="Arial" color="#EFFBBF" size="1"><strong>(required) 
              </strong></font></td>
          </tr>
          <tr> 
            <td align="right" width="149" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>Secret 
              Question :</strong></font></td>
            <td width="307"><font face="Arial" size="2" color="#EFFBBF"> 
              <select name="selquestion" size="1">
                <option value="" <%if qt="" then%> selected <%end if%>>Select 
                Question---------------------</option>
                <%
					call fillcob("question",qt)
				%>
              </select>
              </font><font face="Arial" color="#EFFBBF" size="1"><strong>(required) 
              </strong></font><font face="Arial" size="2" color="#EFFBBF"> </font></td>
          </tr>
          <tr> 
            <td align="right" width="149" valign="top" height="31"><font face="Arial" size="2" color="#EFFBBF"><strong>Answer 
              to Question:</strong></font></td>
            <td width="307" height="31"><font face="Arial" size="2" color="#EFFBBF"> 
              <small> 
              <input type="text" name="txtanswer" maxlength="50"size="24"  value="<%=aw%>">
              </small> </font><font face="Arial" color="#EFFBBF" size="1"><strong>(required) 
              </strong></font></td>
          </tr>
          <tr> 
            <td align="right" width="149" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>First 
              Name:</strong></font></td>
            <td width="307"><font face="Arial" size="2" color="#EFFBBF"> <small> 
              <input type="text" name="txtfirstname" maxlength="50"size="24"  value="<%=fn%>">
              </small> </font></td>
          </tr>
          <tr> 
            <td align="right" width="149" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>Last 
              Name:</strong></font></td>
            <td width="307"><font face="Arial" size="2" color="#EFFBBF"> <small> 
              <input type="text" name="txtlastname" maxlength="50"size="24"  value="<%=ln%>">
              </small> </font></td>
          </tr>
          <tr> 
            <td align="right" width="149" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>Gender:</strong></font></td>
            <td width="307"><font face="Arial" color="#EFFBBF" size="1"> 
              <input type="radio" name="sex" value="mr" <% If gd="mr" then %>checked<% End If %> >
              Mr 
              <input type="radio" name="sex" value="ms" <% If gd="ms" then %>checked<% End If %> >
              Ms </font> </td>
          </tr>
          <tr> 
            <td align="right" width="149" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>Subscribe:</strong></font></td>
            <td width="307"> 
              <input type="radio" name="sub" value="0" <% If sb="0" then %>checked<% End If %>>
              <font face="Arial" color="#EFFBBF" size="1">Yes</font> 
              <input type="radio" name="sub" value="1" <% If sb="1" then %>checked<% End If %> >
              <font face="Arial" color="#EFFBBF" size="1">No</font></td>
          </tr>
          <tr> 
            <td align="right" width="149" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>Country:</strong></font></td>
            <td width="307"> 
              <select name="selstate" size="1">
                <option <% If ct="" then %> selected <% End If %> value="" >Select 
                Country-----------------------</option>
                <%
					call fillcob("country",ct)
				%>
              </select>
              <font face="Arial" color="#EFFBBF" size="1"><strong>(required) </strong></font> 
            </td>
          </tr>
          <tr> 
            <td align="right" width="149" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>Company 
              size :</strong></font></td>
            <td width="307"> 
              <select size="1" name="selsize">
                <option <% If cs="" then %>selected<% End If %> value="" >Select 
                Company Size ----------------</option>
                <%
					call fillcob("csize",cs)
				%>
              </select>
              <font face="Arial" color="#EFFBBF" size="1"><strong>(required) </strong></font> 
            </td>
          </tr>
          <tr> 
            <td align="right" width="149" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>Company 
              Name :</strong></font></td>
            <td width="307"><font face="Arial" size="2" color="#EFFBBF"> 
              <input type="text" size="25"
              name="txtcompanyname"value="<%=cn%>">
              </font><font face="Arial" color="#EFFBBF" size="1"><strong>(required)</strong></font></td>
          </tr>
          <tr> 
            <td align="right" width="149" valign="top"><font face="Arial" size="2" color="#EFFBBF"><strong>Business 
              Type:</strong></font></td>
            <td width="307"> 
              <select size="1" name="seltype">
                <option <% If bt="" then %>selected <% End If %> value="" >Select 
                Business Type-----------------</option>
                <%
					call fillcob("btype",bt)
				%>
              </select>
              <font face="Arial" color="#EFFBBF" size="1"><strong>(required) </strong></font> 
            </td>
          </tr>
          <tr> 
            <td width="149" align="right" valign="top">&nbsp;</td>
            <td valign="bottom" width="307"><font face="Arial" size="2" color="#EFFBBF"> 
              <input type="button" name="Submit" value="Submit" onClick="javascript:dosubmit();">
              <input type="reset" name="Submit2" value="Reset">
              <input type="hidden" name="action_button2" value="">
              </font></td>
        </table>
      </form>

 

</td>
  </tr>
</table>
</body>
</html>