<%
sub pass2
  response.write "<P>Pass 2 tokenvalue="+cstr(tokenvalue)

  set cn=Server.CreateObject("ADODB.Connection")
  cn.open "gl1329","gl1329","FRT46pzyF"
  response.write "<P>Connection created OK"
   set rs=nothing
   response.write "<P>Existence for bid = "+cstr("bidvalue")
        if c=0 then
           response.write ": Found "+cstr(c)+" records. Proceeding with Insert"
           SQLString="INSERT INTO activity (activity_id, a_name, age_group, theme) VALUES ('1','Bugs','Infant','Math')"
           response.write "<P>Ready for Insert with SQLString="+cstr(SQLString)
           cn.execute SQLString,numa
           if numa=1 then
                response.write "<P>Inserted OK numa="+cstr(numa)
           else
                response.write "<P>Insert Failed. Number of records inserted="+cstr(numa)
           end if
           cn.close
           set cn=nothing
           response.write "<P>Terminating normally"
        else
           if c=1 then
                response.write ": Found "+cstr(c)+" records. No insert will be performed. Duplicate key"
           else
                response.write "<P>Found "+cstr(c)+" records. No Insert will be performed. More than one reocrd with this bid found"
           end if
        end if
cn.close

end sub
%>