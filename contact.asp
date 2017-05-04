<!--#include file="database.inc"-->

<%
act_bt=request("txtmail")
straction=request("action_button")
email=request("email")

if straction="save" then
         
         fn = replace(request("name"),"'","''")
         em = replace(request("email"),"'","''")      
		 cn = replace(request("company"),"'","''")
		 phone = replace(request("phone"),"'","''")
         Comment=replace(request("comment"),"'","''")
         				 
		if trim((comment)&"")="" then
		    comment=null
		end if
		
		dat = date()
     
	    Sql = "Insert into Question (Name, Company, Email, Tel, Question) Values "
	    
	    Sql  =  Sql  &  "('"& fn &"', '"& cn &"', '"& em &"', '"& phone &"', '"& Comment &"')"
    
	    Conn.Execute Sql
	    
	    strtemp = "There is enquiry in the Enhance website. Please go to http://www.enhance-tech.com.hk/admin_enquiry.asp?email="&em
	    	    
           Set jmail = Server.CreateObject("JMail.Message")
		'jmail.AddRecipient "sale@enhance-tech.com.hk"
		jmail.From="info@enhance-tech.com.hk"
		jmail.Subject = "There is enquiry from the web."
		jmail.Body = strtemp
		jmail.Send "smtp.bbmail.com.hk"


		response.redirect "enquiry_result.asp"	 
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

<script language="JavaScript">
<!--

function dosubmit()
{
	var errmsg;
	var strmail;
	errmsg="";
	strmail="";

	if(document.updata.name.value=="")
	errmsg = errmsg + "請輸入名稱。\n";

    strmail=document.updata.email.value;
	if(strmail=="" )
		errmsg = errmsg + "請輸入電郵地址。\n";
	else
		if(strmail.search("@")==-1)
			errmsg = errmsg + "請輸入正確電郵地址。\n";

	
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

<!--#include file="menu_bar.js"-->


</HEAD>

</HEAD>
<BODY>

<div class="container">


<!--#include file="menu_nav.inc"-->

   <div class="panel panel-info">

    <div class="panel-heading">

     <table  class="table">

            <tr>

    <td>地址</td>

        <td>東莞樟木頭泰安城一期 1508 室</td></tr>

    	<tr><td>電話:</td>

       <td>+86-769-82052036</td></tr>

         <tr><td>手機:</td> 

         <td>+86-133 8019 4136</td></tr> 

         <tr><td>傳真:</td> 

            <td>+86-769-82052036</td> 

       </tr>
     </table>

</div>

<div class="panel-body">

  <form name="updata" mothed="post" action="contact.asp">

       　<table  class="table">
            <tr>
              <td align="right" width="128" bgcolor="#FFFFFF">
              名稱:</td>
              <td width="448" bgcolor="#FFFFFF">
                    <input type="text" size="53" name="name" value="<%=name%>">(必須)
                    </td>
            </tr>
            <tr>
              <td align="right" width="128" bgcolor="#FFFFFF">公司:</td>
              <td width="448" bgcolor="#FFFFFF">
                    <input type="text" size="53"
              name="company" value="<%=company%>">
                    </td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top" bgcolor="#FFFFFF">電郵
              :</font></td>
              <td width="448" bgcolor="#FFFFFF">
                    <input type="text" size="53"
              name="email" value="<%=email%>">
                     (必須)
</td>
            </tr>
            <tr>
              <td align="right" width="128" bgcolor="#FFFFFF">電話
              :</td>
              <td width="448" bgcolor="#FFFFFF">
                    <input type="text" size="53"
              name="phone" value="<%=phone%>">
                    </td>
            </tr>
            <tr>
              <td align="right" width="128" bgcolor="#FFFFFF" height="115">問題:</td> 
              <td width="448" bgcolor="#FFFFFF" height="115">
				<textarea rows="8" name="Comment" cols="39"></textarea></td>
            </tr>
            <tr>
              <td width="128" align="right" valign="top" bgcolor="#FFFFFF"></td>
              <td valign="bottom" width="448" bgcolor="#FFFFFF">
			  <input type="button" value="提交" onClick="javascript:dosubmit();">
			  <input type="hidden" name="action_button" value=""></td> 
            </tr>
          </table>
		  </form>

    </div>

      </div>



<!--#include file="menu_bottom.inc"-->

</div>
</BODY>
</html>
