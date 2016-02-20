<html>
    <%
       set rs=Server.CreateObject("ADODB.Recordset")
       SQLString="SELECT them FROM activity WHERE age_group ='Allentown'" 
       rs.open SQLString,"DSN=gl1329;UID=gl1329;PWD=FRT46pzyF;"
        response.write "<P>Recordset opened OK"
        response.write "<table border='1'>"
        response.write "<tr><td>Badge id</td><td>Location</td><td>First Name</td><td>Last Name</td></tr>"
           c=0
        response.write "<tr>"
           while NOT rs.EOF
               response.write "<td>"+cstr(rs("theme"))+"</td>"
              c=c+1
              rs.movenext
           wend
           response.write "</tr></table><p>record count="+cstr(c)+"<p>Terminating Normally"
           rs.close
           set rs=nothing
    %>
</html>