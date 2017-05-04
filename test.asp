<html>
<head>
<title>test</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script language="JavaScript">
function doa()
{
 document.form1.select1=enabled;
 alert("AAA");
 //alert(document.form1.select1.selected)
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" action="">
  <input type="radio" name="radiobutton" value="radiobutton" onclick="javascript:doa();">
  <select disabled name="select1">
    <option>aa</option>
    <option>aaaa</option>
    <option>ddddd</option>
  </select>
</form>
</body>
</html>
