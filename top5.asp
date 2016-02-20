<html>
    <%
       set rs=Server.CreateObject("ADODB.Recordset")
       SQLString="SELECT top(5) theme FROM activity WHERE age_group ='Allentown'" 
       rs.open SQLString,"DSN=gl1329;UID=gl1329;PWD=FRT46pzyF;"
        response.write "<P>Recordset opened OK"
        response.write "<table border='1'>"
           c=0
        response.write "<tr>"
           while NOT rs.EOF
               response.write "<td><img src='"+cstr(rs("theme"))+"' alt='test' height='50' width='50'></td>"
              c=c+1
              rs.movenext
           wend
           response.write "</tr></table><p>record count="+cstr(c)+"<p>Terminating Normally"
           rs.close
           set rs=nothing
    %>
</html>