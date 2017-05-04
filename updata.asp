<!--#include file="database.inc"-->
 
<%
 if  session("em")="" or session("id")="" then
        response.redirect "gologin.asp"
    end if
straction=request("action_button") 
strname=replace(request("txtemail"),"'","''")
id=trim(request("chkid"))
if straction="save" then
           strsql="select email from member where email='"& strname &"'  and id<>"& id &""
	'  response.write strsql
	   set acres=conn.execute(strsql)
	   if not acres.eof then
	         response.redirect "updatefailure.asp"
	   else				
				 ps=replace(request("txtpass"),"'","''")
				 qt=replace(request("txtquestion"),"'","''")
				 aw=replace(request("txtanswer") ,"'","''")
				 fn=replace(request("txtfirstname"),"'","''")
				 ln=replace(request("txtlastname") ,"'","''")
				 em=replace(request("txtemail"),"'","''")
				 gd=request("sex")
				 sb=request("sub")
				 ct=replace(request("state"),"'","''")
				 cn=replace(request("txtcompanyname"),"'","''")
				 cs=replace(request("size"),"'","''")
				  bt=replace(request("type"),"'","''")
				  dat=date()
				  straction=trim(request("Action_Button"))
				  sqLcmd= "update  member set Password='"& ps &"',Question='"& qt&"', Answer='"& aw &"',Firstname='"
				  sqlcmd=sqlcmd& fn&"',Lastname='"& ln &"',Sex='"
				  sqlcmd=sqlcmd& gd&"',Country='"& ct &"',CompanyName='"& cn &"',ComanySize='"
				  sqlcmd=sqlcmd& cs &"',BusinessType='"& bt &"',subscribe="& sb &",email='"& strname&"',create_date='"& dat &"' where id= "& id&""
			'	 response.write sqlcmd
			     conn.Execute sqlcmd 
	end if
  %>
  
<body bgcolor="#FFFFFF" link="#0000CC" vlink="#0000CC" alink="#0000CC" >
<table border="0"align="center">
  <%if session("admin")=1 then%>
  <tr> 
    <td colspan="2" height="55" width="500" bgcolor="#006699"><b><i><font color="#ffffff"size="6"> 
      <center>
        <font color="#FFFFFF">Member Information Update Success! </font> 
      </center>
      </font></i></B></td>
  </tr>
  <%else%>
  <tr> 
    <td colspan="2" height="55" width="500" bgcolor="#006699"><b><i><font color="#ffffff"size="6"> 
      <center>
        <font color="#FFFFFF">Your Profile Update Success! </font> 
      </center>
      </font></i></B></td>
  </tr>
  <tr> 
    <td height="30" bgcolor="#FFFFCC"><font color="#666666"></font></td>
  </tr>
  <tr> 
    <td colspan="2" bgcolor="#FFFFCC"> 
      <center>
        <font color="#666666"><a href=show.asp><font size=5><b>Look Your Information</B></font></a></font> 
      </center>
    </td>
  </tr>
  <tr> 
    <td height="10" bgcolor="#FFFFCC"><font color="#666666"></font></td>
  </tr>
  <tr> 
    <td colspan="2" bgcolor="#FFFFCC"> 
      <center>
        <font color="#666666"><a href=updata.asp><font size=5><b>Change Your Information</b></font></a></font> 
      </center>
    </td>
  </tr>
  <%end if%>
  <tr> 
    <td height="15" bgcolor="#FFFFCC"><font color="#666666"></font></td>
  </tr>
</table>
 </body> 
  
 <%
else
  id=trim(request("chkid"))
   if id="" then
         id=session("id")
         sqlcmd="select*from member where id="& session("id")&" "
   else
       sqlcmd="select * from member where id="& id &""
	end if  
 ' response.write sqlcmd
       set rs=conn.execute(sqlcmd)
	  if  not rs.eof then
	      '  name=trim(rs("loginname")&"")
			ps=trim(rs("password")&"")
			qt=replace(rs("question"),"''","'")
			aw=trim(rs("answer")&"")
			fn=trim(rs("firstname")&"")
			ln=trim(rs("lastname")&"")
			gd=trim(rs("sex")&"")
			sb=trim(rs("subscribe")&"")
			em=trim(rs("email")&"")
			ct=replace(rs("country"),"''","'")
			cn=trim(rs("companyname")&"")
        moble=trim(rs("moble")&"")		
			cs=replace(rs("comanysize"),"''","'")		
			bt=rs("businesstype")
			dat=trim(rs("create_date")&"")
		'response.write "country="&ct&"companysize="&cs&"businesstype="&bt
	 end if
     rs.close          
     conn.close 
 %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>Member Profile</title>
<style type="text/css">
 ul {font-family: "Arial", "Helvetica", "sans-serif"; font-size: 11px; color: #000066; font-weight: bold}
 table {  font-family: "Arial", "Helvetica", "sans-serif"; color: #0000FF; font-weight: bold; font-size: 14px}
 picposition {  margin-left: 15px; margin-top: 0px; margin-bottom: 0px}</style>
<base target="_self">

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
	var strpwd;
	
	errmsg="";
	strmail="";

	 strpw=document.updata.txtpass.value;
  if(strpw=="")
	   errmsg="Please Input Password.\n";
  else
		{if(strpw.length<5 )
		errmsg=errmsg+"Password is short! Please input length 5up\n";}			
	if(document.updata.txtquestion.value=="")
	errmsg=errmsg+"Please Select question.\n";
	if(document.updata.txtanswer.value=="")
	errmsg=errmsg+"Please Input answer.\n";
	if(document.updata.txtfirstname.value=="")
	errmsg=errmsg+"Please Input first name.\n";
	if(document.updata.txtlastname.value=="" )
	errmsg=errmsg+"Please Input Last Name.\n";

	strmail=document.updata.txtemail.value;
	if(strmail=="" )
		errmsg=errmsg+"Please Input e-mail address.\n";
	else
		if(strmail.search("@")==-1)
			errmsg=errmsg+"Please Input a right e-mail address.\n";
	if(document.updata.sex.value=="")
	errmsg=errmsg+"Please Input gender.\n";
	
	if(document.updata.txtcompanyname.value=="")
	errmsg=errmsg+"Please Input company Name.\n";
   
	
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

<body bgcolor="#ffffff" >



<form  name="updata" bgcolor='bule' method="post" action="updata.asp">

          
<table border="1"bordercolor="#ffffff" align="center"><center>                 
        <td colspan="2" height="44" bgcolor="#336699"fontcolor="#ffffff"><font color="#cccccc"size="5">
<center><b><i>Edit your profile now!</i></b></center></font></td> 
               <tr>
			   
        <td colspan="2" height="20" bgcolor="#336699"fontcolor="#ffffff"><font color="#cccccc"size="5">
<center>
            <font size="3">Please fill out the form below!</font> 
          </center></font></td> 
</tr>
			    <tr>    
                  
        <td width="149" height="21" bgcolor="#336699"fontcolor="#ffffff"> <font face="Verdana, Geneva, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
          E-Mail <font color="#FF0000" size="2"> </font></font></td>    
                  
        <td width="288"height="21" bgcolor="#ffffcc"fontcolor="#ffffff"> 
          <input type="text" name="txtemail" size="23" maxlength="50" value="<%=em%> " >
          <font face="Verdana, Geneva, Arial, Helvetica, sans-serif" size="1" color="#6666FF">(required)</font></td>    
                </tr>
                <tr >   
                  
        <td width="149" height="21"bgcolor="#336699"fontcolor="#ffffff"> <font face="Verdana, Geneva, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
          Password</font></td>  
                  
        <td width="288" height="21"bgcolor="#ffffcc"fontcolor="#ffffff"> 
          <input type="password" size="23" name="txtpass" maxlength="20" value="<%=ps%>">
          <font color="#FF0000" size="2"> <font face="Verdana, Geneva, Arial, Helvetica, sans-serif" size="1" color="#6666FF">(required)</font></font></td>  
                </tr>  
                
				<tr >   
                  
        <td width="149" height="21"bgcolor="#336699"fontcolor="#ffffff"><font face="Verdana, Geneva, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
          Secret Question </font></td>  
                  
        <td width="288" height="21"bgcolor="#ffffcc"fontcolor="#ffffff"> 
          <select name="txtquestion">
            <option value="" <% if qt="" then%>selected<%end if%>>Select Question 
            </OPTION>
            <option  <%if qt="What's your telephone number?" then%> selected <%end if%> value="What's your telephone number?"> 
            What's your telephone number?</option>
            <option <%if qt="What's your birthday?" then%> selected <%end if%> value="What's your birthday?"> 
            What's your birthday?</option>
            <option  <%if qt="Which country you like most?" then%> selected <%end if%> value="Which country you like most?">Which 
            country you like most?</option>
            <option  <%if qt="What's color you like?" then%> selected <%end if%> value="What's color you like?">What's 
            color you like?</option>
          </select>
          <font color="#6666FF" size="2"> <font face="Verdana, Geneva, Arial, Helvetica, sans-serif" size="1">(required)</font></font></td>
				  <tr >   
                  
        <td width="149" height="21"bgcolor="#336699"fontcolor="#ffffff"><font face="Verdana, Geneva, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
          Answer to Question </font> </td>  
                  
        <td width="288" height="21"bgcolor="#ffffcc"fontcolor="#ffffff"> 
          <input type="text" size="23" name="txtanswer" maxlength="50" value="<%=aw%>">
          <font color="#FF0000" size="2"> <font face="Verdana, Geneva, Arial, Helvetica, sans-serif" size="1" color="#6666FF">(required)</font></font></td>  
                </tr>  
              
				<tr valign="bottom">  
                  
        <td width="149" height="21" bgcolor="#336699"fontcolor="#ffffff"><font face="Verdana, Geneva, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
          First Name </font> </td>  
                  
        <td width="288" height="21"bgcolor="#ffffcc"fontcolor="#ffffff"> 
          <input type="text"  size="23" name="txtfirstname" maxlength="20" value="<%=fn%>">
          <font face="Verdana, Geneva, Arial, Helvetica, sans-serif" size="1" color="#6666FF">(required)</font></td>  
                </tr> 
				<tr valign="bottom">  
                  
        <td width="149" height="21" bgcolor="#336699"fontcolor="#ffffff"><font face="Verdana, Geneva, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
          Last Name </font> </td>  
                  
        <td width="288" height="21"bgcolor="#ffffcc"fontcolor="#ffffff"> 
          <input type="text" size="23" name="txtlastname" maxlength="20" value="<%=ln%>">
          <font color="#FF0000" size="2"> <font face="Verdana, Geneva, Arial, Helvetica, sans-serif" size="1" color="#6666FF">(required)</font></font></td>  
                </tr>          
                <tr valign="bottom">                   
        <td width="149" height="15"bgcolor="#336699"fontcolor="#ffffff"> <font face="Verdana, Geneva, Arial, Helvetica, sans-serif" color="#FFFFFF">Gender 
          </font> </td>  
                  
        <td width="288" height="15"bgcolor="#ffffcc"fontcolor="#0066ff"> 
          <input type="radio"name="sex"value="mr"<% IF gd="mr" THEN %> checked <% end if %>> <font face="Verdana, Geneva, Arial, Helvetica, sans-serif"> 
        <font color="#0066ff">Mr</font></font><input type="radio" name="sex"value="ms"  <% IF gd="ms" THEN %> checked <% end if %>>
          <font face="Verdana, Geneva, Arial, Helvetica, sans-serif"> <font color="#0066ff"> 
          Ms<font color="#6666FF"> </font></font></font>   <font color="#6666FF" size="2"> 
          <font face="Verdana, Geneva, Arial, Helvetica, sans-serif" size="1">
        </font></font> </td>    
                </tr>
 <tr>     
                  
        <td width="149"bgcolor="#336699"fontcolor="#ffffff"><font face="Verdana, Geneva, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
          Moble </font> </td>    
                  
        <td width="288"bgcolor="#ffffcc"fontcolor="#ffffff"> 
          <input type="text" name="moble"  size="23"  value="<%=moble%>">
          <font face="Verdana, Geneva, Arial, Helvetica, sans-serif" size="1" color="#6666FF">(required)</font></td>    
                </tr>  
				 <tr valign="bottom">     
                  
               
        <td width="149" height="21"bgcolor="#336699"fontcolor="#ffffff"><font face="Verdana, Geneva, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
          Subscribe</font> </td>    
                  
        <td width="288" height="27"bgcolor="#ffffcc"fontcolor="#0066ff"> 
          <input type="radio" name="sub" value="0" <%if sb=0 then%> checked<%end if%>>
				 <font color="#0066ff"> Yes</font> <input type="radio" name="sub" value="1" <% if sb=1 then%>checked <%end if%> ><font color="#0066ff"> No </font></td>    
                </tr>
				 <tr>     
                  
        <td width="149"bgcolor="#336699"fontcolor="#ffffff"><font face="Verdana, Geneva, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
          Country </font> </td> 
				
        <td width="288"bgcolor="#ffffcc"fontcolor="#ffffff"> 
          <select NAME="state" size="1" > 
	    <option value="" <% if ct=""then %> selected <% end if %>> Select Country</option>
		 <option <% IF ct="Afghanistan" THEN %> selected <% end if %>>Afghanistan </OPTION>                
        <option <% IF ct="Albania" THEN %> selected <% end if %>>Albania </OPTION>
		<OPTION <% IF ct="American Samoa" THEN %> selected <% end if %>>American Samoa</OPTION>
        <option <% IF ct="Algeria" THEN %> selected <% end if %>>Algeria </OPTION>
        <option <% IF ct="Andorra" THEN %> selected <% end if %>>Andorra </OPTION>
        <option<% IF ct="Angola" THEN %> selected <% end if %> >Angola </OPTION>
        <option <% IF ct="Anguilla" THEN %> selected <% end if %>>Anguilla </OPTION>
        <option <% IF ct="Antarctica" THEN %> selected <% end if %>>Antarctica </OPTION>
        <option <% IF ct="Antigua and Barbuda" THEN %> selected <% end if %>>Antigua and Barbuda </OPTION>
        <option <% IF ct="Argentina" THEN %> selected <% end if %>>Argentina</OPTION>
        <option <% IF ct="Armenia" THEN %> selected <% end if %>>Armenia </OPTION>
        <option <% IF ct="Aruba" THEN %> selected <% end if %>>Aruba </OPTION>
        <option <% IF ct="Australia" THEN %> selected <% end if %>>Australia</OPTION>        
        <option <% IF ct="Azerbaijan" THEN %> selected <% end if %>>Azerbaijan</OPTION> 
		 <OPTION <% IF ct="Alabama" THEN %> selected <% end if %>>Alabama</OPTION>			  
		<OPTION <% IF ct="Alaska" THEN %> selected <% end if %>>Alaska</OPTION>			   
		<OPTION <% IF ct="Arizona" THEN %> selected <% end if %>>Arizona</OPTION>			   
		<OPTION <% IF ct="Arkansas" THEN %> selected <% end if %>>Arkansas</OPTION>			   
		<option <% IF ct="Bahamas" THEN %> selected <% end if %>>Bahamas</OPTION> 			   
        <option  <% IF ct="Bahrain" THEN %> selected <% end if %>>Bahrain</OPTION> 
        <option  <% IF ct="Bangladesh" THEN %> selected <% end if %>>Bangladesh</OPTION> 
        <option  <% IF ct="Barbados" THEN %> selected <% end if %>>Barbados </OPTION>
        <option <% IF ct="Belarus" THEN %> selected <% end if %>>Belarus </OPTION>
        <option  <% IF ct="Belgium" THEN %> selected <% end if %>>Belgium </OPTION>
        <option <% IF ct="Belize" THEN %> selected <% end if %>>Belize </OPTION>
        <option  <% IF ct="Benin" THEN %> selected <% end if %>>Benin </OPTION>
        <option <% IF ct="Bermuda" THEN %> selected <% end if %> >Bermuda </OPTION>
        <option  <% IF ct="Bhutan" THEN %> selected <% end if %>>Bhutan </OPTION>
        <option  <% IF ct="Bolivia" THEN %> selected <% end if %>>Bolivia </OPTION>
        <option <% IF ct="Bosnia and Herzegowina" THEN %> selected <% end if %> >Bosnia and Herzegowina</OPTION> 
        <option <% IF ct="Botswana" THEN %> selected <% end if %>>Botswana </OPTION>
        <option  <% IF ct="Bouvet Island" THEN %> selected <% end if %>>Bouvet Island</OPTION> 
        <option  <% IF ct="Brazil" THEN %> selected <% end if %>>Brazil </OPTION>
        <option <% IF ct="British Indian Ocean Territory" THEN %> selected <% end if %>>British Indian Ocean Territory </OPTION>
        <option  <% IF ct="Brunei Darussala" THEN %> selected <% end if %>>Brunei Darussalam </OPTION>
        <option  <% IF ct="Bulgaria" THEN %> selected <% end if %>>Bulgaria </OPTION>
        <option  <% IF ct="Burkina Faso" THEN %> selected <% end if %>>Burkina Faso </OPTION>
        <option <% IF ct="Burundi" THEN %> selected <% end if %> >Burundi </OPTION>
        <option <% IF ct="Cambodia" THEN %> selected <% end if %> >Cambodia </OPTION>
        <option  <% IF ct="Cameroon" THEN %> selected <% end if %>>Cameroon</OPTION> 
        <option  <% IF ct="Cape Verde" THEN %> selected <% end if %>>Cape Verde </OPTION>
        <option  <% IF ct="Cayman Islands" THEN %> selected <% end if %>>Cayman Islands </OPTION>
        <option  <% IF ct="Central African Republic" THEN %> selected <% end if %>>Central African Republic</OPTION>
        <option  <% IF ct="Chad" THEN %> selected <% end if %>>Chad </OPTION>
        <option <% IF ct="Chile" THEN %> selected <% end if %> >Chile </OPTION>
        <option <% IF ct="China" THEN %> selected <% end if %> >China </OPTION>
        <option  <% IF ct="Christmas Island" THEN %> selected <% end if %>>Christmas Island </OPTION>
        <option  <% IF ct="Cocoa (Keeling) Islands" THEN %> selected <% end if %>>Cocoa (Keeling) Islands </OPTION>
        <option <% IF ct="Colombia" THEN %> selected <% end if %>>Colombia </OPTION>
        <option <% IF ct="Comoros" THEN %> selected <% end if %> >Comoros</OPTION> 
        <option <% IF ct="Congo" THEN %> selected <% end if %> >Congo </OPTION>
        <option <% IF ct="Cook Islands" THEN %> selected <% end if %> >Cook Islands</OPTION>
        <option  <% IF ct="Costa Rica" THEN %> selected <% end if %>>Costa Rica </OPTION>
        <option  <% IF ct="Cote Divoire" THEN %> selected <% end if %>>Cote Divoire </OPTION>
        <option <% IF ct="Croatia (local name: Hrvatska)" THEN %> selected <% end if %> >Croatia (local name: Hrvatska) </OPTION>
        <option <% IF ct="Cuba" THEN %> selected <% end if %> >Cuba </OPTION>
        <option <% IF ct="Cyprus" THEN %> selected <% end if %>>Cyprus </OPTION>
        <option <% IF ct="Czech Republic" THEN %> selected <% end if %> >Czech Republic </OPTION>
		<OPTION <% IF ct="California" THEN %> selected <% end if %>>California</OPTION>			   
		<OPTION <% IF ct="Colorado" THEN %> selected <% end if %>>Colorado</OPTION>			   
		<OPTION <% IF ct="Connecticut" THEN %> selected <% end if %>>Connecticut</OPTION>			   					   
		<OPTION <% IF ct="Delaware" THEN %> selected <% end if %>>Delaware</OPTION>			   
		<OPTION <% IF ct="District of Columbia" THEN %> selected <% end if %>>District of Columbia</OPTION>			   
		<option  <% IF ct="Denmark" THEN %> selected <% end if %>>Denmark</OPTION>			    
        <option  <% IF ct="Djibouti" THEN %> selected <% end if %>>Djibouti</OPTION> 
        <option <% IF ct="Dominica" THEN %> selected <% end if %> >Dominica </OPTION>
        <option <% IF ct="Dominican Republic" THEN %> selected <% end if %> >Dominican Republic</OPTION> 
        <option <% IF ct="East Timor" THEN %> selected <% end if %>>East Timor</OPTION> 
        <option <% IF ct="Ecuador" THEN %> selected <% end if %> >Ecuador </OPTION>
        <option <% IF ct="Egypt" THEN %> selected <% end if %> >Egypt </OPTION>
        <option <% IF ct="El Salvador" THEN %> selected <% end if %> >El Salvador</OPTION> 
        <option <% IF ct="Equatorial Guinea" THEN %> selected <% end if %>>Equatorial Guinea </OPTION>
        <option <% IF ct="Eritrea" THEN %> selected <% end if %>>Eritrea</OPTION>
        <option  <% IF ct="Estonia" THEN %> selected <% end if %>>Estonia </OPTION>
        <option  <% IF ct="Ethiopia" THEN %> selected <% end if %> >Ethiopia </OPTION>
        <option  <% IF ct="Falkland Islands (Malvinas)" THEN %> selected <% end if %>>Falkland Islands (Malvinas)</OPTION> 
        <option  <% IF ct="Faroe Islands" THEN %> selected <% end if %>>Faroe Islands </OPTION>
        <option <% IF ct="Fiji" THEN %> selected <% end if %> >Fiji</OPTION> 
        <option <% IF ct="Finland" THEN %> selected <% end if %>>Finland </OPTION>
        <option  <% IF ct="France, Metropolitan" THEN %> selected <% end if %>>France, Metropolitan </OPTION>
        <option <% IF ct="French Guiana" THEN %> selected <% end if %>>French Guiana</OPTION> 
        <option  <% IF ct="French Polynesia" THEN %> selected <% end if %>>French Polynesia</OPTION> 
        <option  <% IF ct="French Southern Territories" THEN %> selected <% end if %>>French Southern Territories</OPTION> 
		<OPTION <% IF ct="Florida" THEN %> selected <% end if %>>Florida</OPTION>			   
		<OPTION <% IF ct="Georgia" THEN %> selected <% end if %>>Georgia</OPTION>			   
		 <option  <% IF ct="Gabon" THEN %> selected <% end if %>>Gabon</OPTION>			    
        <option  <% IF ct="Gambia" THEN %> selected <% end if %>>Gambia </OPTION>
		<OPTION <% IF ct="Guam" THEN %> selected <% end if %>>Guam</OPTION>
        <option <% IF ct="Georgia" THEN %> selected <% end if %> >Georgia </OPTION>
        <option <% IF ct="Ghana" THEN %> selected <% end if %> >Ghana </OPTION>
        <option <% IF ct="Gibraltar" THEN %> selected <% end if %> >Gibraltar</OPTION> 
        <option <% IF ct="Greece" THEN %> selected <% end if %> >Greece </OPTION>
        <option  <% IF ct="Greenland" THEN %> selected <% end if %>>Greenland</OPTION> 
        <option  <% IF ct="Grenada" THEN %> selected <% end if %>>Grenada</OPTION> 
        <option <% IF ct="Guadeloupe" THEN %> selected <% end if %> >Guadeloupe</OPTION> 
        <option <% IF ct="Guam" THEN %> selected <% end if %> >Guam </OPTION>
        <option <% IF ct="Guatemala" THEN %> selected <% end if %> >Guatemala</OPTION> 
        <option <% IF ct="Guinea" THEN %> selected <% end if %> >Guinea </OPTION>
        <option <% IF ct="Guinea-Bissau" THEN %> selected <% end if %> >Guinea-Bissau </OPTION>
        <option <% IF ct="Guyana" THEN %> selected <% end if %> >Guyana </OPTION>
		 <OPTION <% IF ct="Hawaii" THEN %> selected <% end if %>>Hawaii</OPTION>
        <option <% IF ct="Haiti" THEN %> selected <% end if %> >Haiti </OPTION>
        <option <% IF ct="Heard and Mc Donald Islands" THEN %> selected <% end if %> >Heard and Mc Donald Islands</OPTION> 
        <option  <% IF ct="Honduras" THEN %> selected <% end if %>>Honduras </OPTION>
        <option  <% IF ct="Hong Kong" THEN %> selected <% end if %> >Hong Kong </OPTION>
        <option  <% IF ct="Hungary" THEN %> selected <% end if %>>Hungary </OPTION>
        <option <% IF ct="Iceland" THEN %> selected <% end if %> >Iceland </OPTION>
		<OPTION <% IF ct="Idaho" THEN %> selected <% end if %>>Idaho</OPTION>
		<OPTION <% IF ct="Illinois" THEN %> selected <% end if %>>Illinois</OPTION>
		<OPTION <% IF ct="Indiana" THEN %> selected <% end if %>>Indiana</OPTION>			   
		<OPTION <% IF ct="Iowa" THEN %> selected <% end if %>>Iowa</OPTION>			  			   
        <option <% IF ct="Indonesia" THEN %> selected <% end if %> >Indonesia </OPTION>
        <option <% IF ct="Iran (Islamic Republic of)" THEN %> selected <% end if %> >Iran (Islamic Republic of) </OPTION>
        <option  <% IF ct="Iraq" THEN %> selected <% end if %>>Iraq </OPTION>
        <option <% IF ct="Ireland" THEN %> selected <% end if %> >Ireland </OPTION>
        <option <% IF ct="Israel" THEN %> selected <% end if %> >Israel </OPTION>
        <option <% IF ct="Italy" THEN %> selected <% end if %> >Italy </OPTION>
        <option <% IF ct="Jamaica" THEN %> selected <% end if %> >Jamaica </OPTION>
        <option  <% IF ct="Japan" THEN %> selected <% end if %>>Japan </OPTION>
        <option  <% IF ct="Jordan" THEN %> selected <% end if %>>Jordan</OPTION> 
		<OPTION <% IF ct="Kansas" THEN %> selected <% end if %>>Kansas</OPTION>			  
		 <OPTION <% IF ct="Kentucky" THEN %> selected <% end if %>>Kentucky</OPTION>			   
		 <option <% IF ct="Kazakhstan" THEN %> selected <% end if %> >Kazakhstan </OPTION>
        <option  <% IF ct="Kenya" THEN %> selected <% end if %>>Kenya </OPTION>
        <option <% IF ct="Kiribati" THEN %> selected <% end if %> >Kiribati </OPTION> 
        <option <% IF ct="Korea, Republic of" THEN %> selected <% end if %>>Korea, Republic of </OPTION>
        <option <% IF ct="Kuwait" THEN %> selected <% end if %> >Kuwait </OPTION>
        <option <% IF ct="Kyrgyzstan" THEN %> selected <% end if %> >Kyrgyzstan</OPTION> 
        <option <% IF ct="Lao Peoples Democratic Republic" THEN %> selected <% end if %>>Lao Peoples Democratic Republic </OPTION>
        <option <% IF ct="Latvia" THEN %> selected <% end if %> >Latvia </OPTION>
        <option <% IF ct="Lebanon" THEN %> selected <% end if %> >Lebanon</OPTION>
        <option <% IF ct="Lesotho" THEN %> selected <% end if %> >Lesotho </OPTION>
        <option <% IF ct="Liberia" THEN %> selected <% end if %> >Liberia </OPTION>
        <option <% IF ct="Libyan Arab Jamahiriya" THEN %> selected <% end if %> >Libyan Arab Jamahiriya </OPTION>
        <option <% IF ct="Liechtenstein" THEN %> selected <% end if %> >Liechtenstein </OPTION>
        <option <% IF ct="Lithuania" THEN %> selected <% end if %>>Lithuania </OPTION>
        <option <% IF ct="Luxembourg" THEN %> selected <% end if %> >Luxembourg </OPTION>			   
		<OPTION <% IF ct="Louisiana" THEN %> selected <% end if %>>Louisiana</OPTION>			  
		<OPTION <% IF ct="Maine" THEN %> selected <% end if %>>Maine</OPTION>			   
		<OPTION <% IF ct="Maryland" THEN %> selected <% end if %>>Maryland</OPTION>			   
		<OPTION <% IF ct="Massachusetts" THEN %> selected <% end if %>>Massachusetts</OPTION>			   
		<OPTION <% IF ct="Michigan" THEN %> selected <% end if %>>Michigan</OPTION>			   
		<OPTION <% IF ct="Minnesota" THEN %> selected <% end if %>>Minnesota</OPTION>			   
		<OPTION <% IF ct="Mississippi" THEN %> selected <% end if %>>Mississippi</OPTION>			   
		 <OPTION <% IF ct="Missouri" THEN %> selected <% end if %>>Missouri</OPTION>			   
		<OPTION <% IF ct="Maldives" THEN %> selected <% end if %> >Maldives </OPTION>
        <option <% IF ct="Mali" THEN %> selected <% end if %> >Mali </OPTION>
        <option <% IF ct="Malta" THEN %> selected <% end if %> >Malta </OPTION>
        <option <% IF ct="Marshall Islands" THEN %> selected <% end if %>>Marshall Islands </OPTION>
        <option <% IF ct="Martinique" THEN %> selected <% end if %> >Martinique </OPTION>
        <option <% IF ct="Mauritania" THEN %> selected <% end if %>>Mauritania </OPTION>
        <option <% IF ct="Mauritius" THEN %> selected <% end if %>>Mauritius</OPTION> 
        <option <% IF ct="Mayotte" THEN %> selected <% end if %> >Mayotte </OPTION>
        <option <% IF ct="Mexico" THEN %> selected <% end if %> >Mexico </OPTION>
        <option <% IF ct="Micronesia, Federated States of" THEN %> selected <% end if %> >Micronesia, Federated States of </OPTION>
        <option <% IF ct="Moldova, Republic of" THEN %> selected <% end if %> >Moldova, Republic of </OPTION>
        <option  <% IF ct="Monaco" THEN %> selected <% end if %>>Monaco </OPTION>
        <option <% IF ct="Mongolia" THEN %> selected <% end if %> >Mongolia </OPTION>
        <option  <% IF ct="Montserrat" THEN %> selected <% end if %>>Montserrat</OPTION> 
        <option <% IF ct="Morocco" THEN %> selected <% end if %> >Morocco </OPTION>
        <option <% IF ct="Mozambique" THEN %> selected <% end if %> >Mozambique</OPTION> 
        <option <% IF ct="Myanmar" THEN %> selected <% end if %> >Myanmar </OPTION>
        <option <% IF ct="Namibia" THEN %> selected <% end if %> >Namibia</OPTION> 
        <option <% IF ct="Nauru" THEN %> selected <% end if %> >Nauru </OPTION>
        <option  <% IF ct="Nepal" THEN %> selected <% end if %>>Nepal </OPTION>
        <option  <% IF ct="New Caledonia" THEN %> selected <% end if %>>New Caledonia </OPTION>
        <option  <% IF ct="New Zealand" THEN %> selected <% end if %>>New Zealand </OPTION>
        <option <% IF ct="Nicaragua" THEN %> selected <% end if %> >Nicaragua </OPTION>
        <option <% IF ct="Niger" THEN %> selected <% end if %> >Niger </OPTION> 
		<OPTION <% IF ct="Northern Mariana Islands" THEN %> selected <% end if %>>Northern Mariana Islands</OPTION>
        <option <% IF ct="Nigeria" THEN %> selected <% end if %> >Nigeria </OPTION>
        <option <% IF ct="Niue" THEN %> selected <% end if %> >Niue </OPTION>
        <option <% IF ct="Norfolk Island" THEN %> selected <% end if %> >Norfolk Island </OPTION>
        <option <% IF ct="Northern Mariana Islands" THEN %> selected <% end if %> >Northern Mariana Islands </OPTION>
        <option <% IF ct="Norway" THEN %> selected <% end if %> >Norway</OPTION>			   
		<OPTION <% IF ct="Nebraska" THEN %> selected <% end if %>>Nebraska</OPTION>			   
		<OPTION <% IF ct="Nevada" THEN %> selected <% end if %>>Nevada</OPTION>			   
		 <OPTION <% IF ct="New Hampshire" THEN %> selected <% end if %>>New Hampshire</OPTION>			  
		<OPTION <% IF ct="New Jersey" THEN %> selected <% end if %>>New Jersey</OPTION>			   
		 <OPTION <% IF ct="New Mexico" THEN %> selected <% end if %>>New Mexico</OPTION>			  
		 <OPTION <% IF ct="New York" THEN %> selected <% end if %>>New York</OPTION>			  
		<OPTION <% IF ct="North Carolina" THEN %> selected <% end if %>>North Carolina</OPTION>			   
		<OPTION <% IF ct="North Dakota" THEN %> selected <% end if %>>North Dakota</OPTION>
		<OPTION <% IF ct="Ohio" THEN %> selected <% end if %>>Ohio</OPTION>
		<OPTION <% IF ct="Oklahoma" THEN %> selected <% end if %>>Oklahoma</OPTION>			   
		 <OPTION <% IF ct="Oregon" THEN %> selected <% end if %>>Oregon</OPTION>			  
		 <option <% IF ct="Oma" THEN %> selected <% end if %> >Oman </OPTION>
        <option  <% IF ct="Pakistan" THEN %> selected <% end if %>>Pakistan</OPTION> 
        <option  <% IF ct="Palau" THEN %> selected <% end if %>>Palau </OPTION>
		<OPTION <% IF ct="Puerto Rico" THEN %> selected <% end if %>>Puerto Rico</OPTION>
        <option <% IF ct="Panama" THEN %> selected <% end if %> >Panama </OPTION>
        <option  <% IF ct="Papua New Guinea" THEN %> selected <% end if %>>Papua New Guinea</OPTION> 
        <option  <% IF ct="Paraguay" THEN %> selected <% end if %>>Paraguay</OPTION> 
        <option <% IF ct="Peru" THEN %> selected <% end if %> >Peru </OPTION>
        <option  <% IF ct="Philippines" THEN %> selected <% end if %>>Philippines</OPTION> 
        <option  <% IF ct="Pitcairn" THEN %> selected <% end if %>>Pitcairn</OPTION> 
        <option <% IF ct="Polan" THEN %> selected <% end if %> >Poland </OPTION>
        <option  <% IF ct="Portugal" THEN %> selected <% end if %>>Portugal </OPTION>
        <option <% IF ct="Puerto Rico" THEN %> selected <% end if %>>Puerto Rico</OPTION> 
		<OPTION <% IF ct="Palau" THEN %> selected <% end if %>>Palau</OPTION>			   
		<OPTION <% IF ct="Pennsylvania" THEN %> selected <% end if %>>Pennsylvania</OPTION>
		<option <% IF ct="Qatar" THEN %> selected <% end if %> >Qatar</OPTION>		   
		<OPTION <% IF ct="Rhode Island" THEN %> selected <% end if %>>Rhode Island</OPTION>			   
		<option <% IF ct="Reunion" THEN %> selected <% end if %>>Reunion </OPTION>
        <option <% IF ct="Romania" THEN %> selected <% end if %> >Romania </OPTION>
        <option  <% IF ct="Russian Federation" THEN %> selected <% end if %>>Russian Federation</OPTION> 
        <option <% IF ct="Rwanda" THEN %> selected <% end if %> >Rwanda </OPTION>
        <option  <% IF ct="Saint Kitts and Nevis" THEN %> selected <% end if %>>Saint Kitts and Nevis </OPTION>
        <option  <% IF ct="Saint Lucia" THEN %> selected <% end if %>>Saint Lucia </OPTION>
        <option  <% IF ct="Saint Vincent and the Grenadines" THEN %> selected <% end if %>>Saint Vincent and the Grenadines </OPTION>
        <option  <% IF ct="Samoa" THEN %> selected <% end if %>>Samoa </OPTION>
        <option  <% IF ct="San Marino" THEN %> selected <% end if %>>San Marino</OPTION> 
        <option  <% IF ct="Sao Tome and Principe" THEN %> selected <% end if %>>Sao Tome and Principe</OPTION> 
        <option  <% IF ct="Saudi Arabia" THEN %> selected <% end if %>>Saudi Arabia </OPTION>
        <option  <% IF ct="Senegal" THEN %> selected <% end if %>>Senegal</OPTION> 
        <option  <% IF ct="Seychelles" THEN %> selected <% end if %>>Seychelles </OPTION>
        <option  <% IF ct="Sierra Leone" THEN %> selected <% end if %>>Sierra Leone</OPTION> 
        <option  <% IF ct="Singapore" THEN %> selected <% end if %>>Singapore </OPTION>
        <option <% IF ct="Slovakia (Slovak Republic)" THEN %> selected <% end if %> >Slovakia (Slovak Republic)</OPTION> 
        <option <% IF ct="Slovenia" THEN %> selected <% end if %>>Slovenia </OPTION>
        <option  <% IF ct="Solomon Islands" THEN %> selected <% end if %>>Solomon Islands </OPTION>
        <option <% IF ct="Somalia" THEN %> selected <% end if %> >Somalia </OPTION>
        <option <% IF ct="South Africa" THEN %> selected <% end if %> >South Africa </OPTION>
        
        <option <% IF ct="Spain" THEN %> selected <% end if %> >Spain </OPTION>
        <option <% IF ct="Sri Lanka" THEN %> selected <% end if %> >Sri Lanka</OPTION> 
        <option  <% IF ct="St. Helena" THEN %> selected <% end if %>>St. Helena </OPTION>
        <option  <% IF ct="St. Pierre and Miquelon" THEN %> selected <% end if %>>St. Pierre and Miquelon</OPTION> 
        <option  <% IF ct="Sudan" THEN %> selected <% end if %>>Sudan </OPTION>
        <option <% IF ct="Suriname" THEN %> selected <% end if %>>Suriname </OPTION>
        <option <% IF ct="Svalbard and Jan Mayen Islands" THEN %> selected <% end if %> >Svalbard and Jan Mayen Islands </OPTION>
        <option <% IF ct="Swaziland" THEN %> selected <% end if %> >Swaziland </OPTION>
        <option <% IF ct="Sweden" THEN %> selected <% end if %> >Sweden</OPTION> 
        <option  <% IF ct="Switzerland" THEN %> selected <% end if %>>Switzerland </OPTION>
        <option  <% IF ct="Syrian Arab Republic" THEN %> selected <% end if %>>Syrian Arab Republic </OPTION>			   
		 <OPTION <% IF ct="South Carolina" THEN %> selected <% end if %>>South Carolina</OPTION>			   
		 <OPTION <% IF ct="Tennessee" THEN %> selected <% end if %> >Tennessee</OPTION>			  
	    <OPTION <% IF ct="Texas" THEN %> selected <% end if %>>Texas</OPTION>
		<option <% IF ct="Taiwan" THEN %> selected <% end if %> >Taiwan</OPTION> 
        <option <% IF ct="Tajikistan" THEN %> selected <% end if %> >Tajikistan </OPTION>
        <option <% IF ct="Tanzania, United Republic of" THEN %> selected <% end if %> >Tanzania, United Republic of </OPTION>
        <option <% IF ct="Thailand" THEN %> selected <% end if %> >Thailand</OPTION> 
        <option <% IF ct="Togo" THEN %> selected <% end if %> >Togo</OPTION> 
        <option <% IF ct="Tokelau" THEN %> selected <% end if %> >Tokelau</OPTION> 
        <option <% IF ct="Tonga" THEN %> selected <% end if %>>Tonga</OPTION> 
        <option <% IF ct="Trinidad and Tobago" THEN %> selected <% end if %> >Trinidad and Tobago</OPTION> 
        <option <% IF ct="Tunisia" THEN %> selected <% end if %> >Tunisia</OPTION> 
        <option  <% IF ct="Turkey" THEN %> selected <% end if %>>Turkey</OPTION> 
        <option <% IF ct="Turkmenistan" THEN %> selected <% end if %> >Turkmenistan </OPTION>
        <option <% IF ct="Turks and Caicos Islands" THEN %> selected <% end if %> >Turks and Caicos Islands</OPTION> 
        <option <% IF ct="Tuvalu" THEN %> selected <% end if %> >Tuvalu</OPTION>
        <option <% IF ct="Uganda" THEN %> selected <% end if %> >Uganda</OPTION> 
        <option <% IF ct="Ukraine" THEN %> selected <% end if %> >Ukraine </OPTION>
        <option <% IF ct="United Arab Emirates" THEN %> selected <% end if %> >United Arab Emirates </OPTION>
        <option  <% IF ct="Uruguay" THEN %> selected <% end if %>>Uruguay </OPTION> 
		<OPTION <% IF ct="Utah" THEN %> selected <% end if %>>Utah</OPTION>
        <option <% IF ct="Uzbekistan" THEN %> selected <% end if %> >Uzbekistan</OPTION> 
        <option <% IF ct="Vanuatu" THEN %> selected <% end if %> >Vanuatu</OPTION> 
        <option <% IF ct="Vatican City State (Holy See)" THEN %> selected <% end if %> >Vatican City State (Holy See)</OPTION>
        <option <% IF ct="Venezuela" THEN %> selected <% end if %>>Venezuela </OPTION>
        <option <% IF ct="Viet Nam" THEN %> selected <% end if %> >Viet Nam </OPTION>
        <option <% IF ct="Virgin Islands (British)" THEN %> selected <% end if %> >Virgin Islands (British)</OPTION> 
        <option <% IF ct="Virgin Islands (U.S.)" THEN %> selected <% end if %> >Virgin Islands (U.S.)</OPTION> 		  
		<OPTION <% IF ct="Virginia" THEN %> selected <% end if %>>Virginia</OPTION>
		 <OPTION <% IF ct="Virgin Island" THEN %> selected <% end if %>>Virgin Island</OPTION>			  
		 <OPTION<% IF ct="Vermont" THEN %> selected <% end if %>>Vermont</OPTION>			  
		 <OPTION<% IF ct="Washington" THEN %> selected <% end if %>>Washington</OPTION>			   
		<OPTION <% IF ct="West Virginia" THEN %> selected <% end if %>>West Virginia</OPTION>			  
		  <OPTION <% IF ct="Wisconsin" THEN %> selected <% end if %>>Wisconsin</OPTION>			   
		<OPTION <% IF ct="Wyoming" THEN %> selected <% end if %>>Wyoming</OPTION>			 
		<option <% IF ct="Wallisw and Futuna Islands" THEN %> selected <% end if %> >Wallisw and Futuna Islands</OPTION> 
        <option <% IF ct="Western Sahara" THEN %> selected <% end if %> >Western Sahara</OPTION> 
        <option <% IF ct="Yeman" THEN %> selected <% end if %> >Yeman </OPTION>
        <option <% IF ct="Yugoslavia" THEN %> selected <% end if %> >Yugoslavia </OPTION>
        <option  <% IF ct="Zaire" THEN %> selected <% end if %>>Zaire</OPTION> 
        <option <% IF ct="Zambia" THEN %> selected <% end if %> >Zambia</OPTION> 
        <option <% IF ct="Zimbabwe" THEN %> selected <% end if %> >Zimbabwe</OPTION>                  
						
					   </select>    
	  </td>   
                </tr>   
				    
                <tr>     
                  
        <td width="149"bgcolor="#336699"fontcolor="#ffffff"><font face="Verdana, Geneva, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
          Company Name </font> </td>    
                  
        <td width="288"bgcolor="#ffffcc"fontcolor="#ffffff"> 
          <input type="text" name="txtcompanyname"  size="23"  value="<%=cn%>">
          <font face="Verdana, Geneva, Arial, Helvetica, sans-serif" size="1" color="#6666FF">(required)</font></td>    
                </tr>   
             
            
                <tr>     
                  
        <td width="149"bgcolor="#336699"fontcolor="#ffffff"><font face="Verdana, Geneva, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
          Company Size</font></td>   
                  
        <td width="288"bgcolor="#ffffcc"fontcolor="#ffffff"> 
          <select size="1" name="size" >
				       <option value=""<%if cs=""then %> selected<% end if%>  >Select Company Size</option>				   
				        <option <% IF cs="under10" THEN %> selected <% end if %> >under10</option>
					   <option <% IF cs="11~25" THEN %> selected <% end if %>>11~25</option>					  
						 <option <% IF cs="26~50" THEN %> selected <% end if %>>26~50</option>
						  <option <% IF cs="51~100" THEN %> selected <% end if %>>51~100</option> 
						  <option <% IF cs="101~500" THEN %> selected <% end if %> >101~500</option>
						  <option <% IF cs="501~1000" THEN %> selected <% end if %> >501~1000</option>
						  <option <% IF cs="1000up" THEN %> selected <% end if %> >1000up</option>
						  </select></td>
				  </tr>                              
                <tr>    
                  
        <td width="149"bgcolor="#336699"fontcolor="#ffffff"><font face="Verdana, Geneva, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
          Business Type </font> </td>   
                  
        <td width="288"bgcolor="#ffffcc"> 
          <select size="1" name="type" >
				       <option value=""<%if bt=""then%> selected <%end if%>>Select Business Type</option>
				       <option <% IF bt="Personal" THEN %> selected <% end if %> >Personal</option>
				       <option  <% IF bt="Importer" THEN %> selected <% end if %> >Importer</option>
					   <option <% IF bt="Wholesale" THEN %> selected <% end if %> >Wholesale</option>
					    <option <% IF bt="Promotional" THEN %> selected <% end if %> >Promotional</option>
				 </select></td>   
                 
             
                </tr> </center>
				 
                <tr>    
                  
      <td colspan="2"bgcolor="#ffffff"> 
        <div align="center" style="width: 287; height: 25">    
                      <input type="button" name="Submit" value="Submit" onClick="javascript:dosubmit();">
					  <input type="reset" name="Submit2" value="Reset"> 
					  <input type="hidden" name="action_button" value=""> 
					   <input type="hidden" name="chkid" value="<%=id%>"> 
				
                    </div>  
                  </td> 
				   
 
 
 
 
</table></form>
</body>
</html>
<%
end if
%>



