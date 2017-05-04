<% 
	Set SConn = Server.CreateObject("ADODB.Connection")
	SConn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & server.mappath("datasearch/mydata.mdb")
 %>
