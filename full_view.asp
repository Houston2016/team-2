<html>
<body>
            <%
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql_string="SELECT * FROM activity ORDER BY activity_id ASC"

            response.write "<p>SQL--->"+sql_string+"<---<p>"
            rs.open sql_string, "DSN=gl1329;UID=gl1329;PWD=FRT46pzyF;"
            response.write "<p>Opened sql_string = " + sql_string + " OK."
            
	    response.write "<p><table border='1'><tr>"
                    for i = 0 to rs.fields.count - 1
                    	response.write "<td align='center'>"+cstr(rs(i).Name)+"</td>"
                    next
            response.write "</tr>"
                row_count=0
                while not rs.EOF
                	row_count=row_count+1
                	response.write "<tr> "
                    		for i = 0 to rs.fields.count - 1
                    			response.write "<td align='right'>"+cstr(rs(i))+"</td>"
                    		next
                    	response.write "</tr>"
                rs.MoveNext
                wend
                response.write "<tr><td colspan='6' align='left'>" + cstr(row_count) + " rows in the the table.</td></tr>"
                response.write "</table><p>Row Count="+cstr(row_count)
            response.write "<p>Normal Termination "+cstr(now)
            rs.Close
            set rs=nothing
            %>
</body>
</html>