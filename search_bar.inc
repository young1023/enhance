<!--- search bar start ---->
<script language="JavaScript">
<!--
function dosearch(p_arg)
{
	var strmsg="";
	
	switch(p_arg){
		case "model":
			if(document.search.txtmodel.value=="")
				strmsg="You must input a word to search.\n";
			break;
		case "category":
			if(document.search.selcategory.value=="-1")
			strmsg="You must select category to search!\n";
			break;
		case "exprice":
			if(document.search.selexprice.value=="-1" )
			strmsg="You must select a export price to search!\n";
			break;
		default:
	}
	if(strmsg!="")
	alert(strmsg);
	else
	{
		document.search.action_button.value=p_arg;
		document.search.submit();
	}
}
function dologin()
{
	if(document.mainlogin.txtemail.value=="" || document.mainlogin.txtpassword.value=="" )
		alert("Please input login email address and password!");
	else
		document.mainlogin.submit();
}

//-->
</script>
      <form name="search" action="searchresult.asp" method="post">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td ><IMG height=1 alt="" 
      src="images/space.gif" width=1 border=0></td>
            <td width="18%"><b><font size="-2"   face="Arial, Helvetica, sans-serif">Model:</font></b>
              <input type="text" name="txtmodel" size="8" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 8pt" value="" maxlength="8">
            <a href="javascript:dosearch('model');"><img src="images/demo/bt-next.jpg" width="14" height="14" border="0"></a></font></b></td>
            <td width="3%"><b><font size="-2"   face="Arial, Helvetica, sans-serif">Price:</font></b></td>
            <td width="20%"><b><font size="-2"   face="Arial, Helvetica, sans-serif"> 
              <select name="selexprice" style="BACKGROUND-COLOR: #f8f8f8; BORDER-BOTTOM: #9a9999 1px solid; BORDER-LEFT: #9a9999 1px solid; BORDER-RIGHT: #9a9999 1px solid; BORDER-TOP: #9a9999 1px solid; FONT-SIZE: 8pt">
              <option selected value="-1"  value="">Please Select Price </option>
					<% 
						strsql="select * from exportprice order by id"
						set acres=conn.execute(strsql)
						do while not acres.eof
					 %>
                    <option value="<%= trim(acres("id")&"") %>"><%= trim(acres("exportprice")&"") %></option>
					<% 
							acres.movenext
						loop
						set acres=nothing
					
					 %>
			  
            </select>
            <a href="javascript:dosearch('exprice');"><img src="images/demo/bt-next.jpg" width="14" height="14" border="0"></a></font></b></td>
			<input type="hidden" name="action_button" value=""> 
        </tr>
      </table>
	</form>
  
<!---search bar end ----->