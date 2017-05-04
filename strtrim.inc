<script language="JavaScript">
<!--
function strtrim(p_arg)
{
	
	var i,j,p_char,retval;
	retval="";
	if (p_arg!="" || p_arg!=null )
	{
		for (i=0 ;i<=p_arg.length-1;i++)
		{
			p_char=p_arg.charAt(i);
			if (p_char!=" ")
			break;
		}
		
		if (i!=p_arg.length-1)
		{
			
			j=p_arg.length-1;
			while (j>=0)
			{
				p_char="";
				p_char=p_arg.charAt(j);
				if (p_char!=" ")break;
				j--;
			}
			if(j!=0)
			retval=p_arg.substr(i,j+1);
		}
		else
			retval=p_arg;
	}
	//alert(p_arg);
		return retval;	
}
//-->
</script>