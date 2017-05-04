
<!--#include file="database.inc"-->
<%
sqlcmd="select*from member where id="& session("id")&""

set rs=conn.execute(sqlcmd) 
if not rs.eof then 
     
      pw=rs("password")      
      qt=rs("question")
      aw=rs("answer")
      fn=rs("firstname")
      ln=rs("lastname")
      gd=rs("sex")
      em=rs("email")
      ct=rs("country")
      cn=rs("companyname")
      cs=rs("ComanySize")
      bt=rs("businesstype")
	  date1=rs("create_date")
'response.write qt
end if
%>
<body bgcolor="#FFFFFF" link="#0000CC" vlink="#0000CC" alink="#0000CC" >
<table border="1" bordercolor="#ffffff" align="center" width="410">
<tr>
               
    <td colspan="2" height="44" width="400" bgcolor="#006699"><font color="#ffffff"size="6"> 
      <center>
        <b><I> Your Profile</i></B> 
      </center></font>
	  </td> 
                </tr> 
				

     <tr> 
    <td bgcolor="#006699" width="151"><font color="#FFFFFF">E_mail </font> 
    <td  width="243" bgcolor="#FFFFCC"><font color="#000000"><%=em%> 
      </font> 
  
  <tr> 
    <td bgcolor="#006699" width="151"><font color="#FFFFFF">Password </font> 
    <td  width="243" bgcolor="#FFFFCC"><font color="#000000"><%=pw%> 
      </font> 
 <%if qt<>""then%>
  <tr> 
    <td bgcolor="#006699" width="151"><font color="#FFFFFF">Secret Question </font> 
    <td width="243" bgcolor="#FFFFCC"><font color="#000000"><%=qt%>
      </font>
	  <%else%> 
	  <tr> 
    <td bgcolor="#006699" width="151"><font color="#FFFFFF">Secret Question </font> 
    <td width="243" bgcolor="#FFFFCC"><font color="#000000">Empty
      </font>
	  <%end if%>
  <%if aw<>""then%>
  <tr>   
    <td bgcolor="#006699" width="151"><font color="#FFFFFF">Answer to Question 
      </font> 
    <td  width="243" bgcolor="#FFFFCC"><font color="#000000"><%=aw%> 
      </font> 
  <% end if%> 
  <tr> 
    <td bgcolor="#006699" width="151"><font color="#FFFFFF">First Name </font> 
    <td  width="243" bgcolor="#FFFFCC"><font color="#000000"><%=fn%> 
      </font> 

  <tr> 
    <td bgcolor="#006699" width="151"><font color="#FFFFFF">Last Name </font> 
    <td  width="243" bgcolor="#FFFFCC"><font color="#000000"><%=ln%> 
      </font> 
  
  <tr> 
    <td bgcolor="#006699" width="151"><font color="#FFFFFF">Gender </font> 
    <td  width="243" bgcolor="#FFFFCC"><font color="#000000"><%=gd%> 
      </font>   
  <%if ct<>""then%>
  <tr> 
    <td bgcolor="#006699" width="151"><font color="#FFFFFF">Country </font> 
    <td  width="243" bgcolor="#FFFFCC"><font color="#000000"><%=ct%> 
      </font> 
  <% end if%>  <%if cs<>""then%>
  <tr> 
    <td bgcolor="#006699" width="151"><font color="#FFFFFF">Company Size </font> 
    <td  width="243" bgcolor="#FFFFCC"><font color="#000000"><%=cs%> 
      </font> <% end if%>
  
  <tr> 
    <td bgcolor="#006699" width="151"><font color="#FFFFFF">Company Name </font> 
    <td  width="243" bgcolor="#FFFFCC"><font color="#000000"><%=cn %> 
      </font> 
    <%if bt<>""then%>
  <tr> 
    <td bgcolor="#006699" width="151"><font color="#FFFFFF">Business Type </font> 
    <td  width="243" bgcolor="#FFFFCC"><font color="#000000"><%=bt%> 
      </font> 
  <% end if%>
  <tr> 
    <td bgcolor="#006699" width="151"><font color="#FFFFFF">Register date </font> 
    <td  width="243" bgcolor="#FFFFCC"><font color="#000000"><%=date1%> 
      </font> 
  

    
<tr> 

 <td colspan="2" width="400"> <center>     
   <a href=updata.asp><b><font size=3 color="#800000">Change Information</font></b></a></center></td> 
 
</table>   
 
</body> 
 
 
 
 
 
