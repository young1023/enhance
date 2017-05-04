<!--#include file="database.inc"-->
<%
  strsql="select a_id from acclist"
    set acres=conn.Execute(strsql)
	    do while not acres.eof
	        carno=trim(acres("A_id"))
	       sql="select* from accmain where U_no="& carno &""
		   set ares=conn.execute(sql)
		   if ares.eof then
		       delsql="deletE * from acclist where a_id="& carno &""
			   set strdesql=conn.execute(delsql)
			   set strdesql=nothing
			   set delsql=nothing
			end if
			set ares=nothing
			set sql=nothing
	 acres.movenext
	loop
    set acres=nothing
    set strsql=nothing
response.redirect "admin_order.asp"		    
		    
%>