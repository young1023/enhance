<%
condition1=trim(request("hidselect"))
session("condition2")=condition1
%>
<!--#include file="database.inc"-->
<html>
<head>
<title>popSelect Member</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script language="javascript">
<!--
var email,name;
 name="";
 email="";
function dook()
{
name=document.frmsearch.hidselect1.value;
if (name=="group")
  email=document.frmsearch.txtgroupname.value;
if (name=="country")
   email=document.frmsearch.txtcountry.value;
if (name=="businesstype")
  email=document.frmsearch.txttype.value;
if (name=="companysize")
   email=document.frmsearch.txtcomsize.value; 
if (name=="alluser")
   email="To send E_mail for all member"; 
if (name=="")
  alert("Please Click Radio to Select");
 else
  {if (email=="")
      alert("Please Select Value");
    else
     { 
	 window.opener.document.frmsend.txtto.value =email;
	 window.close() ;}
  }
}
function dopop1()
{
name="group";
document.frmsearch.hidselect.value=name;
document.frmsearch.submit();
}
function dopop2()
{
name="country";
document.frmsearch.hidselect.value=name;
document.frmsearch.submit();
}
function dopop3()
{
name="businesstype";
document.frmsearch.hidselect.value=name;
document.frmsearch.submit();
}
function dopop4()
{
name="companysize";
document.frmsearch.hidselect.value=name;
document.frmsearch.submit();
}
function dopop5()
{
name="alluser";
document.frmsearch.hidselect.value=name;
document.frmsearch.submit();
}
-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#0000CC" vlink="#0000CC" alink="#0000CC">
<form name="frmsearch" >
  <p><font color="#FF0000"><b>Please Click Radio ,And Then Select Right Select 
    To Send Email </b></font></p>
    
  <table width="46%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="47%" bgcolor="#006699" height="26"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif">Group:</font></td>
      <td width="6%" bgcolor="#006699" height="26"> 
        <div align="center"> <font color="#FFFFFF"> 
          <input type="radio" name="radpop" value="group"  <%if condition1="group" then%>checked <%end if%>onclick="javascript:dopop1();" >
          </font></div>
      </td>
      <td width="47%" bgcolor="#FFFFCC" height="26"> <font color="#000000" face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
        <select  <%if condition1="group" then%> enabled <% else %> disabled <%end if%>  name="txtgroupname">
          <option value="" selected><font color="#000000" face="Verdana,Geneva,Arial,Helvetica,sans-serif">Please 
          Select Group</font></option>
          <%		
		   strsql="select distinct(group_name) from user_group order by group_name"
							set acres=conn.execute(strsql)
							do while not acres.eof		 %>
          <option value="<%=trim(acres("group_name")&"")%>"><font color="#000000" face="Verdana,Geneva,Arial,Helvetica,sans-serif"><%=trim(acres("group_name")&"")%></font></option>
          <%acres.movenext
						loop
			 %>
        </select>
        </font> </td>
    </tr>
    <tr> 
      <td width="47%" bgcolor="#006699" height="31"><font color="#FFFFFF" face="Verdana,Geneva,Arial,Helvetica,sans-serif">Country:</font></td>
      <td width="6%" bgcolor="#006699" height="31"> 
        <div align="center"> <font color="#FFFFFF"> 
          <input type="radio" name="radpop" value="country"<%if condition1="country" then%>checked <%end if%> onclick="javascript:dopop2();">
          </font></div>
      </td>
      <td width="47%" bgcolor="#FFFFCC" height="31"> <font color="#000000" face="Verdana,Geneva,Arial,Helvetica,sans-serif"> 
        <select <%if condition1="country" then%> enabled <% else %> disabled <%end if%>  name="txtcountry">
          <option value="" selected>Please Select County</option>
          <% 
			   strsql="select distinct (country) from member order by country"
						set acres=conn.execute(strsql)
						do while not acres.eof %>
          <option value="<%=trim(acres("country")&"")%>"><%=trim(acres("country")&"")%></option>
          <%acres.movenext
				     loop
				   '  set acres=nothing
		
				 %>
        </select>
        </font> </td>
    </tr>
    <tr> 
      <td width="47%" bgcolor="#006699" height="31"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#FFFFFF">Business 
        Type:</font></td>
      <td width="6%" bgcolor="#006699" height="31"> 
        <div align="center"> <font color="#FFFFFF"> 
          <input type="radio" name="radpop" value="businesstype"   <%if condition1="businesstype" then%>checked <%end if%> onclick="javascript:dopop3();" >
          </font></div>
      </td>
      <td width="47%" bgcolor="#FFFFCC" height="31"> <font color="#000000"> 
        <select <%if condition1="businesstype" then%> enabled <% else %> disabled <%end if%>  name="txttype">
          <option value="" selected>Please Select Business Type</option>
          <option value="Importer">Importer</option>
          <option value="Wholesale">Wholesale</option>
          <option value="Promotional">Promotional</option>
          <option value="Personal">Personal</option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td width="47%" bgcolor="#006699" height="37"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#FFFFFF">Company 
        Size:</font></td>
      <td width="6%" bgcolor="#006699" height="37"> 
        <div align="center"> <font color="#FFFFFF"> 
          <input type="radio" name="radpop" value="comanysize"   <%if condition1="companysize" then%>checked <%end if%> onclick="javascript:dopop4();">
          </font></div>
      </td>
      <td width="47%" bgcolor="#FFFFCC" height="37"> <font color="#000000"> 
        <select <%if condition1="companysize" then%> enabled <% else %> disabled <%end if%>  name="txtcomsize">
         <option  value="" selected>Select Company Size </OPTION>
					   <option   >under 10</option>
					   <option >11~25</option>
					  	 <option >26~50</option>
						  <option>51~100</option> 
						  <option >101~500</option>
						  <option >501~1000</option>
						  <option >1000up</option>
						  </select>
        </font></td>
    </tr>
	    <tr> 
      <td width="47%" bgcolor="#006699" height="37"><font face="Verdana,Geneva,Arial,Helvetica,sans-serif" color="#FFFFFF">All 
        User :</font></td>
      <td width="6%" bgcolor="#006699" height="37"> 
        <div align="center"> <font color="#FFFFFF"> 
          <input type="radio" name="radpop" value="alluser"   <%if condition1="alluser" then%>checked <%end if%> onclick="javascript:dopop5();">
          </font></div>
      </td>
      <td width="47%" bgcolor="#FFFFCC" height="37"> <font color="#000000"> 
        <%if condition1="alluser"  then%>
        <font color="#00CC66">E_mail will send to All Member</font><font color="#00FF00"> 
        </font> 
        <%end if%>
        </font></td>
    </tr>
    <tr> 
      <td colspan="3" bgcolor="#FFFFFF"> 
        <input type="button" name="cmdsearck" value="OK" onclick="javascript:dook();">
        <font color="#000000"> 
        <input type="hidden" name="hidselect" value="">
        <input type="hidden" name="hidselect1" value="<%=condition1%>">
        </font> </td>
    </tr>
  </table>
  <p>&nbsp;</p>
  </form>
<p>&nbsp;</p>
</body>
</html>
