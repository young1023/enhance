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
<meta name="keywords" content="GC-heat, EOC , Hasco, strack, DME, progressive, Temperature Controller, band heaters, coil heaters, date inserts, Thermocouple, sensor, shot counter, hydraulic cylinder, mold and die components and functional elements, cooling system, industrial heating, vega;wema;hotset;opitz;GC-Heater;�[����;�o����;�o����;�ūױ��;���q��;�[����;�[����;�o����;�o��;�ű��c;�ɰ�;�ɰ��ަ������q;">
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
	errmsg = errmsg + "�п�J�W�١C\n";

    strmail=document.updata.email.value;
	if(strmail=="" )
		errmsg = errmsg + "�п�J�q�l�a�}�C\n";
	else
		if(strmail.search("@")==-1)
			errmsg = errmsg + "�п�J���T�q�l�a�}�C\n";

	
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

    <td>�a�}</td>

        <td>�F��̤��Y���w���@�� 1508 ��</td></tr>

    	<tr><td>�q��:</td>

       <td>+86-769-82052036</td></tr>

         <tr><td>���:</td> 

         <td>+86-133 8019 4136</td></tr> 

         <tr><td>�ǯu:</td> 

            <td>+86-769-82052036</td> 

       </tr>
     </table>

</div>

<div class="panel-body">

  <form name="updata" mothed="post" action="contact.asp">

       �@<table  class="table">
            <tr>
              <td align="right" width="128" bgcolor="#FFFFFF">
              �W��:</td>
              <td width="448" bgcolor="#FFFFFF">
                    <input type="text" size="53" name="name" value="<%=name%>">(����)
                    </td>
            </tr>
            <tr>
              <td align="right" width="128" bgcolor="#FFFFFF">���q:</td>
              <td width="448" bgcolor="#FFFFFF">
                    <input type="text" size="53"
              name="company" value="<%=company%>">
                    </td>
            </tr>
            <tr>
              <td align="right" width="128" valign="top" bgcolor="#FFFFFF">�q�l
              :</font></td>
              <td width="448" bgcolor="#FFFFFF">
                    <input type="text" size="53"
              name="email" value="<%=email%>">
                     (����)
</td>
            </tr>
            <tr>
              <td align="right" width="128" bgcolor="#FFFFFF">�q��
              :</td>
              <td width="448" bgcolor="#FFFFFF">
                    <input type="text" size="53"
              name="phone" value="<%=phone%>">
                    </td>
            </tr>
            <tr>
              <td align="right" width="128" bgcolor="#FFFFFF" height="115">���D:</td> 
              <td width="448" bgcolor="#FFFFFF" height="115">
				<textarea rows="8" name="Comment" cols="39"></textarea></td>
            </tr>
            <tr>
              <td width="128" align="right" valign="top" bgcolor="#FFFFFF"></td>
              <td valign="bottom" width="448" bgcolor="#FFFFFF">
			  <input type="button" value="����" onClick="javascript:dosubmit();">
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
