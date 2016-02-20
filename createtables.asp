<html>
    <head>
        <img src="mfrm.jpg" alt="mfrm">
    </head>

<body>
<% '*************************** functions, subs, then main 
sub buildProgress (cnnew)
dim create_string
on error resume next
   create_string="CREATE TABLE progress ("
   create_string=create_string +"sequence integer NOT NULL,"
   create_string=create_string +"uid varchar(10) NOT NULL,"
   create_string=create_string +"time_stamp datetime,"
   create_string=create_string +"rating varchar(1) NOT NULL)"
   cnnew.execute create_string 
   if noerrors(cnnew, "Task: Create new areaManager table") then
         Response.write "<br>4. Created new areaManager table OK"
   else 
         Response.write "<br>4. Create new areaManager table task failed *************************<br>"
end if
end sub

sub dropProgress (cnnew)
on error resume next
   cnnew.execute "DROP TABLE progress", numa
   if noerrors(cnnew, "Task: drop progress table") then
      Response.write "<br>3. Dropped old progress table OK"
   else 
      Response.write "<br>3. Unable to drop progress table. Task Failed ***********************<br>" 
   end if
   buildProgress (cnnew)
end sub

sub buildUser (cnnew)
dim create_string
on error resume next
   create_string="CREATE TABLE progress ("
   create_string=create_string +"sequence integer NOT NULL,"
   create_string=create_string +"uid varchar(10) NOT NULL,"
   create_string=create_string +"time_stamp datetime,"
   create_string=create_string +"rating varchar(1) NOT NULL)"
   cnnew.execute create_string 
   if noerrors(cnnew, "Task: Create new areaManager table") then
         Response.write "<br>4. Created new areaManager table OK"
   else 
         Response.write "<br>4. Create new areaManager table task failed *************************<br>"
end if
end sub

sub dropareaManager (cnnew)
on error resume next
   cnnew.execute "DROP TABLE areaManager", numa
   if noerrors(cnnew, "Task: drop areaManager table") then
      Response.write "<br>3. Dropped old areaManager table OK"
   else 
      Response.write "<br>3. Unable to drop areaManager table. Task Failed ***********************<br>" 
   end if
   buildareaManager (cnnew)
end sub


Function noerrors (cn , task)
If Err <> 0 Then
    If cn.Errors.Count = 0 Then

    Else
         for i = 0 to cn.Errors.Count - 1
              response.write "<p>"
              response.write cn.errors(i)
         next
    End If
    noerrors= False
Else
    noerrors = True
End if
End Function

'*************************** main ****************************************

dim cm,cnold,cnnew
dim sumbal, rsold, Insert_String,create_string,numa,numnew,deval
dim fdsn, fuid,fpwd
sumbal=0
numnew=0
on error resume next

'************ Change the three lines below to your credentials

fdsn="gl1329"
fuid="gl1329"
fpwd="FRT46pzyF"

'*********** open the user requested user database

set cnnew = Server.CreateObject("ADODB.Connection")
cnnew.open fdsn,fuid,fpwd

if noerrors (cnnew, "Task: Opening database") then '******** top test for user database
          Response.write "<br>2. Opened your "
          Response.write  fdns
          Response.write " database OK"

          call dropareaManager(cnnew) '****** drop, then create the areaManager table 

          Response.write "<p><table border="+chr(34)+"1"+chr(34)+">"
          Response.write "<tr><td>Store Number</td><td>Store Name</td><td>District</td><td>Area Manager</td><td></tr>"
     

cnnew.close

else '**************************** couldn't open the user's database -- bail out!

       Response.write "<p>2. Open task failed on database: "
       Response.write cStr(fdsn)
       Response.write "<p>3. Drop current areaManager table NOT attempted"
       Response.write "<p>4. Create areaManager table NOT attempted"
       Response.write "<p>5. Drop current je table NOT attempted"
       Response.write "<p>6. Create je table NOT attempted"

end if 

'**************************** back to HTML for the finish

%>
<p>
<center><b>END OF THE CREATETABLES FILE</center>
<p>
</body>
</html>