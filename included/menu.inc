<table width="180" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td valign="top"> 
      <table align=center border=0 cellpadding=0 cellspacing=0 width=180>
        <tr> 
          <td height="28" align=center><h4><font color="#FF005F">Development Guideline</font></td>
        </tr>
        <tr> 
          <td bgcolor=#FFCC33 height=28>&nbsp;&nbsp;<a href="user.asp">Procedure Guide</a></td>
        </tr>
<tr> 
          <td bgcolor=#FFCC33 height=28>&nbsp;&nbsp;&nbsp;&nbsp;<a href="sa_project.asp">Project</a></td>
        </tr>
<tr> 
          <td bgcolor=#FFCC33 height=28>&nbsp;&nbsp;&nbsp;&nbsp;<a href="sa_specification.asp">Specification</a></td>
        </tr>
<tr> 
          <td bgcolor=#FFCC33 height=28>&nbsp;&nbsp;&nbsp;&nbsp;<a href="sa_testing.asp">Testing & Modifcation</a></td>
        </tr>

        <tr> 
          <td bgcolor=#FFCC33 height=28>&nbsp;&nbsp;<a href="user_info.asp">Update 
            Information</a></td>
        </tr>
        <tr> 
          <td bgcolor=#FFCC33 height=28>&nbsp;&nbsp;<a href="/eastzone/index.asp">Daily Work Log</a></td>
        </tr>
		<%if session("shell_power")="8" or session("shell_power")="3" then
          response.write "<tr>"
          response.write "<td bgcolor='#FFCC33' height=28>&nbsp;&nbsp;<a href='sa_member.asp'>Membership</a></td>"
          response.write "</tr>"
          response.write "<tr>"
          response.write "<td bgcolor='#FFCC33' height=28>&nbsp;&nbsp;<a href='listbill.asp'>Form Management</a></td>"
          response.write "</tr>"
          response.write "<tr>"
          response.write "<td bgcolor='#FFCC33' height=28>&nbsp;&nbsp;<a href='sa_news.asp'>News Management</a></td>"
          response.write "</tr>"
          response.write "<tr>"
          response.write "<td bgcolor='#FFCC33' height=28>&nbsp;&nbsp;<a href='sa_download.asp'>Download Database</a></td>"
          response.write "</tr>"
		  end if
		%>
        <tr>
          <td bgcolor=#FFCC33 height=28>&nbsp;&nbsp;<a href="Logout.asp">Logout</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>