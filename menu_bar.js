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
