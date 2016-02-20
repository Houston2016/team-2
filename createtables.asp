<html>
<body>
<b>test</b>
<% '*************************** functions, subs, then main 
sub buildProgress (cnnew)
dim create_string
on error resume next
   create_string="CREATE TABLE progress ("
   create_string=create_string +"sequence integer NOT NULL,"
   create_string=create_string +"uid varchar(10) NOT NULL,"
   create_string=create_string +"time_stamp datetime,"
   create_string=create_string +"rating varchar(1) NOT NULL,"
    create_string=create_string +"feedback varchar(250) NOT NULL)"
   cnnew.execute create_string 
   if noerrors(cnnew, "Task: Create new progress table") then
         Response.write "<br>4. Created new progress table OK"
   else 
         Response.write "<br>4. Create new progress table task failed *************************<br>"
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
   create_string="CREATE TABLE user1 ("
   create_string=create_string +"uid integer NOT NULL,"
   create_string=create_string +"first_name varchar(250) NOT NULL,"
   create_string=create_string +"last_name varchar(250) NOT NULL)"
   cnnew.execute create_string 
   if noerrors(cnnew, "Task: Create new user table") then
         Response.write "<br>4. Created new user table OK"
   else 
         Response.write "<br>4. Create new user table task failed *************************<br>"
end if
end sub

sub dropUser (cnnew)
on error resume next
   cnnew.execute "DROP TABLE user1", numa
   if noerrors(cnnew, "Task: drop user table") then
      Response.write "<br>3. Dropped old user table OK"
   else 
      Response.write "<br>3. Unable to drop user table. Task Failed ***********************<br>" 
   end if
   buildUser (cnnew)
end sub

sub buildActivity (cnnew)
dim create_string
on error resume next
   create_string="CREATE TABLE activity ("
   create_string=create_string +"activity_id integer NOT NULL,"
   create_string=create_string +"a_name varchar(250) NOT NULL,"
   create_string=create_string +"age_group varchar(250) NOT NULL,"
   create_string=create_string +"theme varchar(250) NOT NULL)"
   cnnew.execute create_string 
   if noerrors(cnnew, "Task: Create new activity table") then
         Response.write "<br>4. Created new progress table OK"
   else 
         Response.write "<br>4. Create new activity table task failed *************************<br>"
end if
end sub

sub dropActivity (cnnew)
on error resume next
   cnnew.execute "DROP TABLE activity", numa
   if noerrors(cnnew, "Task: drop activity table") then
      Response.write "<br>3. Dropped old activity table OK"
   else 
      Response.write "<br>3. Unable to drop activity table. Task Failed ***********************<br>" 
   end if
   buildActivity (cnnew)
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

          call dropProgress(cnnew) '****** drop, then create the areaManager table 
          call dropActivity(cnnew)
          call dropUser(cnnew)
     

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