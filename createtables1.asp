<html>

<% '*************************** functions, subs, then main 
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
         Response.write "<br>4. Created new activity table OK"
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
   buildactivity (cnnew)
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

          call dropactivity(cnnew) '****** drop, then create the activity table 

          Response.write "<p><table border="+chr(34)+"1"+chr(34)+">"
          Response.write "<tr><td>Activity_ID</td><td>a_name</td><td>age_group</td><td>theme</td><td></tr>"

            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001000','Houston-Whse','Houston','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001005','Nasa','Houston-Southeast','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001007','Rice','Houston-South','Stokman, Jonathon');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001016','Wayside','Houston-South','Guerra, Alejandro');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001017','Sugar Land','Houston-Southwest','Dykes, Jeremy K.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001018','Galleria','Houston-Southcentral','Buchner, Carl');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001019','Copperwood','Houston-West','Campitelli, Joseph A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001020','Edloe','Houston-South','Stokman, Jonathon');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001022','Woodlands','Houston-North','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001023','Fry Road','Houston-West','Bucklew, David R.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001024','Southwest','Houston-Southcentral','Buchner, Carl');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001025','Pasadena','Houston-Southeast','Ross, James C.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001027','El Campo','Houston-Southwest','Mendez, Jesus A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001029','North Freeway','Houston-Central','Bradshaw, Christopher W.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001030','Cutten','Houston-Northwest','Nilsson, Robert N.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001031','Fondren','Houston-WestCentral','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001032','Northwest','Houston-WestCentral','Morales, Fernando');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001033','Baytown','Houston-Northeast','Panganiban, Juan');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001035','El Dorado','Houston-Southeast','Wayland, Sarah E.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001036','Conroe','Houston-North','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001037','Galleria Mall - Houston','Houston-Southcentral','Leturno, Scott');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001038','Pearland','Houston-South','Hendricks, Olivia L.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001039','Riverpark','Houston-Southwest','Rizvi, Mohsin H.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001040','West Oaks Clearance','Houston-SouthwestCentral','Grosenbacher, Eric M.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001041','Buffalo Speedway','Houston-South','Hendricks, Olivia L.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001042','1960 Clearance','Houston-Northwest','Nilsson, Robert N.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001043','Galveston','Houston-Southeast','Behrmann, Joshua A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001044','Humblewood','Houston-Northeast','Kane, Blake A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001045','Tomball','Houston-Northwest','Jackson, Kyle T.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001046','Cypress','Houston-West','Chavez Jr., Jose M.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001047','League City','Houston-Southeast','Escoubas, Jacob W.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001048','Westheimer @ Shepherd','Houston-Southcentral','Seraphin, Dominique G.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001049','Heights','Houston-Central','Alston, Michael');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001050','Hedwig Village','Houston-WestCentral','Blake, John T.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001051','Kirby','Houston-South','Stokman, Jonathon');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001052','First Colony','Houston-Southwest','Dykes, Jeremy K.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001053','Conroe Marketplace','Houston-North','Cavallaro, Vincent');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001054','Atascocita','Houston-Northeast','Kane, Blake A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001056','Vision Park','Houston-North','Ortiz, Eric F.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001057','Spring (AmerMatt)','Houston-Northwest','Williamson, Virgil');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001058','Echo/Bunker (Amer/Matt)','Houston-WestCentral','Madkins, Raquel Y.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001060','Fry Road (AmerMatt)','Houston-West','Bucklew, David R.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001061','Baytown (AmerMatt)','Houston-Northeast','Panganiban, Juan');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001062','Dairy Ashford (AmerMatt)','Houston-SouthwestCentral','Townsend, Crystal N.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001064','Fondren (AmerMatt)','Houston-Central','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001065','Humble (AmerMatt)','Houston-Northeast','Holmes, Spencer');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001068','Sugar Land (AmerMatt)','Houston-Southwest','Rizvi, Mohsin H.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001070','Eldridge (AmerMatt)','Houston-WestCentral','Morales, Fernando');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001071','Sawyer Heights / I10 (AmerMatt)','Houston-Central','Alston, Michael');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001073','Spring Cypress Commons','Houston-Northwest','Williamson, Virgil');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001074','Bunker Hill','Houston-WestCentral','Wilson, Benjamin B.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001077','West Oaks North (Pro 503)','Houston-West','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001078','Spring Cypress (Pro)','Houston-West','Chavez Jr., Jose M.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001081','Rosenberg (Pro)','Houston-Southwest','Mendez, Jesus A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001082','League City Marketplace','Houston-Southeast','Behrmann, Joshua A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001083','Fairway Center','Houston-Southeast','Escoubas, Jacob W.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001084','Meyerland','Houston-Southcentral','Leturno, Scott');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001085','Westchase','Houston-WestCentral','Wilson, Benjamin B.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001086','Bella Terra (aka Cinco Ranch)','Houston-SouthwestCentral','Kopchinsky, Edward A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001087','Lake Jackson Super Center','Houston-South','Guerra, Alejandro');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001088','Bissonett','Houston-SouthwestCentral','Jones Jr., Carl M.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001089','Willow Chase Shopping Mall','Houston-Northwest','Ellis, Emma');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001090','Mariner Village Shopping Center','Houston-WestCentral','Blake, John T.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001094','Marque','Houston-WestCentral','Madkins, Raquel Y.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001095','Huntsville','Houston-North','Cavallaro, Vincent');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001096','Wallisville','Houston-Northeast','Kane, Blake A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001098','Lakewood Forest','Houston-Northwest','Williamson, Virgil');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001099','Gallery SuperCenter','Houston-Central','Bradshaw, Christopher W.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001100','Atascocita East','Houston-Northeast','Medel, Dalila A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001101','Nasa (MG)','Houston-Southeast','Wayland, Sarah E.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001102','Sugar Land 90','Houston-Southwest','Rizvi, Mohsin H.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001103','Champion (MG)','Houston-Northwest','Nilsson, Robert N.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001104','Humble (MG)','Houston-Northeast','Panganiban, Juan');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001105','Woodlands (MG)','Houston-North','Ortiz, Eric F.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001106','Conroe (MG)','Houston-North','Cavallaro, Vincent');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001108','Tomball Crossing (MG)','Houston-Northwest','Jackson, Kyle T.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001109','Woodlands - Pinecroft  (MG)','Houston-North','Cholewa, Michael P.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001110','Cypress TX (MG)','Houston-West','Campitelli, Joseph A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001111','Shops at Stone Park (MG)','Houston-Northeast','Medel, Dalila A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001112','New West Oaks (MG)','Houston-SouthwestCentral','Townsend, Crystal N.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001113','West Loop (MG)','Houston-Southcentral','Seraphin, Dominique G.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001114','Copperfield (MG)','Houston-Northwest','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001115','Katy Westgate (MG)','Houston-West','Muras, Jason A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001117','Dunvale (MG)','Houston-WestCentral','Wilson, Benjamin B.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001118','Spring Valley (MG)','Houston-WestCentral','Madkins, Raquel Y.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001119','Cinco Ranch (MG)','Houston-SouthwestCentral','Jones Jr., Carl M.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001120','Rice Village (MG)','Houston-South','Stokman, Jonathon');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001121','Missouri City (MG)','Houston-Southwest','Jones, Nicholas B.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001122','Sugarland / HWY 59 (MG)','Houston-Southwest','Dykes, Jeremy K.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001123','Pearland (MG)','Houston-South','Hendricks, Olivia L.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001124','Rosenberg (MG)','Houston-Southwest','Mendez, Jesus A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001125','League City (MG)','Houston-Southeast','Behrmann, Joshua A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001126','Fountains (MG)','Houston-SouthwestCentral','Kopchinsky, Edward A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001128','Sienna Plantation','Houston-Southwest','Jones, Nicholas B.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001130','Pearland West','Houston-South','Guerra, Alejandro');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001131','Cypresswood','Houston-Northwest','Ellis, Emma');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001132','Willowbrook (Relo 0f ST0128)','Houston-Northwest','Ellis, Emma');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001134','Midtown Twenty Six','Houston-Central','Alston, Michael');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001135','Magnolia','Houston-North','Cholewa, Michael P.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001137','Kingwood','Houston-Northeast','Holmes, Spencer');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001138','Yale Street Market','Houston-Central','Thacker, Christopher P.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001139','Louetta & Cutten','Houston-Northwest','Williamson, Virgil');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001140','Brenham','Houston-West','Chavez Jr., Jose M.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001141','Mason Road','Houston-SouthwestCentral','Grosenbacher, Eric M.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001142','Woodway Arch','Houston-WestCentral','Blake, John T.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001143','Meyerland II','Houston-Southcentral','Leturno, Scott');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001144','Baybrook Square','Houston-Southeast','Escoubas, Jacob W.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001145','West Gray','Houston-Central','Thacker, Christopher P.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001146','Mount Belvieu','Houston-Northeast','Medel, Dalila A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001147','Long Meadow Farms','Houston-SouthwestCentral','Kopchinsky, Edward A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001148','Pearland Parkway','Houston-South','Guerra, Alejandro');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001149','Kings Crossing','Houston-Northeast','Holmes, Spencer');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001150','Tanglewood Court','Houston-WestCentral','Blake, John T.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001151','Weslayan','Houston-Southcentral','Seraphin, Dominique G.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001152','Westheimer & Montrose','Houston-Central','Thacker, Christopher P.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001153','Baybrook Passage','Houston-Southeast','Wayland, Sarah E.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001154','I-45','Houston-North','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001157','Post Oak Blvd','Houston-Southcentral','Leturno, Scott');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001160','Memorial Thicket','Houston-WestCentral','Morales, Fernando');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001164','Baytown Marketplace','Houston-Northeast','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001165','Missouri City South','Houston-Southwest','Jones, Nicholas B.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001166','Oxford Plaza','Houston-Central','Bradshaw, Christopher W.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001167','Pearland Town Center','Houston-South','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001168','Montgomery Crossing','Houston-North','Cavallaro, Vincent');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001169','South Mason','Houston-SouthwestCentral','Grosenbacher, Eric M.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001170','Lake Woodlands','Houston-North','Cholewa, Michael P.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001171','Bella Terra','Houston-SouthwestCentral','Jones Jr., Carl M.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001172','Katy (Pro 501)','Houston-West','Bucklew, David R.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001173','Copperfield (Pro)','Houston-West','Campitelli, Joseph A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001174','Meadows Marketplace','Houston-SouthwestCentral','Jones Jr., Carl M.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001175','Westgate','Houston-West','Bucklew, David R.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001176','Tomball','Houston-Northwest','Jackson, Kyle T.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001177','Cypress Mill Plaza','Houston-West','Chavez Jr., Jose M.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001178','Magnolia','Houston-North','Cholewa, Michael P.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001179','Riceland Pavillion','Houston-Northeast','Panganiban, Juan');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001180','West Crossing','Houston-SouthwestCentral','Townsend, Crystal N.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001181','River Park','Houston-Southwest','Rizvi, Mohsin H.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001182','Friendswood','Houston-Southeast','Ross, James C.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001183','Nasa Parkway','Houston-Southeast','Ross, James C.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001184','Spring','Houston-North','Ortiz, Eric F.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001185','Willowbrook Commons','Houston-Northwest','Nilsson, Robert N.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001186','Kirkwood Square','Houston-SouthwestCentral','Kopchinsky, Edward A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001187','Shepherd at 610','Houston-Central','Bradshaw, Christopher W.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001188','Westheimer & Montrose','Houston-Central','Thacker, Christopher P.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001189','Westheimer & Helena','Houston-Central','Alston, Michael');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001190','Post Oak @ San Felipe','Houston-Southcentral','Seraphin, Dominique G.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001191','Memorial Retail Center','Houston-WestCentral','Morales, Fernando');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001192','Market @ Town Center','Houston-Southwest','Dykes, Jeremy K.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001193','Grand Crossing II','Houston-West','Muras, Jason A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001194','Sienna Parkway','Houston-Southwest','Jones, Nicholas B.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('098001','HLSR - Tent Patio','Houston-MCS','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('098101','HLSR - Tent AF','Houston-MCS','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401001','Montgomery Crossing MP','Houston MP-North','Cholewa, Michael P.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401002','South Mason MP','Houston MP-South','Jones Jr., Carl M.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401003','Lake Woodlands MP','Houston MP-North','Cholewa, Michael P.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401004','Bella Terra MP','Houston MP-South','Jones Jr., Carl M.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401005','Katy (Pro 501) MP','Houston MP-South','Jones Jr., Carl M.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401006','Copperfield (Pro) MP','Houston MP-North','Campitelli, Joseph A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401007','Meadows Marketplace MP','Houston MP-South','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401008','Westgate MP','Houston MP-South','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401009','Tomball MP','Houston MP-North','Campitelli, Joseph A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401010','Cypress Mill Plaza MP','Houston MP-South','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401011','Magnolia MP','Houston MP-North','Cholewa, Michael P.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401012','Riceland Pavillion MP','Houston MP-North','Morales, Fernando');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401013','West Crossing MP','Houston MP-South','Townsend, Crystal N.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401014','River Park MP','Houston MP-South','Townsend, Crystal N.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401015','Friendswood MP','Houston MP-North','Morales, Fernando');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401016','Nasa Parkway MP','Houston MP-North','Morales, Fernando');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401017','Spring MP','Houston-Southeast','Cholewa, Michael P.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401018','Willowbrook Commons MP','Houston MP-North','Campitelli, Joseph A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401019','Kirkwood Square MP','Houston MP-South','Townsend, Crystal N.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401020','Shepherd at 610 MP','Houston MP-North','Campitelli, Joseph A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401021','Westheimer & Montrose MP','Houston MP-South','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401022','Westheimer & Helena MP','Houston MP-South','Jones Jr., Carl M.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401023','Post oak @ San Felipe MP','Houston MP-South','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001127','Houston MG Warehouse (MG)','Houston','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401027','Market @ Town Center MP','Houston MP-South','Townsend, Crystal N.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401029','Sienna Parkway MP','Houston MP-South','Townsend, Crystal N.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401028','Grand Crossing MP','Houston MP-South','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('401026','Memorial Retail Center MP','Houston MP-South','Jones Jr., Carl M.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001155','Southwest II','Houston-Southcentral','Buchner, Carl');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001158','Uptown Crossing','Houston-Southcentral','Buchner, Carl');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001156','Grand Crossing  ','Houston-West','Muras, Jason A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001159','Greens Landing','Houston-Central','Bradshaw, Christopher W.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001162','River Pointe','Houston-Southcentral','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001163','Woodpark','Houston-North','Ortiz, Eric F.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001161','Four Corners','Houston-Southwest','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001195','Fairfield Town Center','Houston','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001196','Oak Forest','Houston','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('001197','Park West Center','Houston','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007005','Buckhead','Atlanta-Central','Grise, Jessica L.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007020','Conyers','Atlanta-Central','Hutton, Matthew B.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007027','Lithonia','Atlanta-Central','Grise, Jessica L.');"

            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007035','Chamblee','Atlanta-Central','Grise, Jessica L.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007037','Conyers Commons','Atlanta-Central','Beckom, Kristen E.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007042','Decatur (Emory)','Atlanta-Central','Hutton, Matthew B.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007052','Buckhead Square','Atlanta-Central','Stack, Rebecca A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007065','Buckhead Crossing','Atlanta-Central','Stack, Rebecca A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007067','Brookhaven Plaza','Atlanta-Central','Stack, Rebecca A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007075','Ansley Park (MG)','Atlanta-Central','Grise, Jessica L.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007087','Conyers Plaza (MG)','Atlanta-Central','Hutton, Matthew B.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007102','Buckhead Georgia (MX)','Atlanta-Central','Stack, Rebecca A.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007112','North Druid','Atlanta-Central','Beckom, Kristen E.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007121','Covington GA','Atlanta-Central','Beckom, Kristen E.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007123','Conyers II','Atlanta-Central','Hutton, Matthew B.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007127','Conyers III','Atlanta-Central','Beckom, Kristen E.');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('150001','Halfmoon','Albany','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('150002','Queensbury','Albany','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('150003','Guilderland','Albany','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('150004','Glenville','Albany','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('150005','Niskayuna','Albany','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('150006','Clifton Park','Albany','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('150007','Colonie','Albany','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('150008','Rensselaer','Albany','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('150009','Hannaford','Albany','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('150010','Latham','Albany','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('150011','Saratoga Sp','Albany','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151001','Airport Center','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151002','South Mall','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151003','Valley Plaza Shopping Center','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151004','Whitehall North','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151005','Whitehall South','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151006','Easton','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151007','Bethlehem','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151008','Trexler Mall','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151009','Reading','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151010','Easton','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151011','Hamburg','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151012','Exeter','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151013','Richland Marketplace','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151014','Quakertown','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151015','5th and Elizabeth Ave','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151016','Trexlertown','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151017','Wyomissing','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('151018','Phillipsburg','Allentown','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
            Insert_String = "INSERT INTO activity (activity_id,a_name,age_group,theme) VALUES ('007131','Lavista Northlake Center','Atlanta-Central','NULL');"
            cnnew.execute Insert_String '****** add the row to the new table
    dim update_string
    update_string = "UPDATE activity SET theme = 'http://assets-s3.usmagazine.com/uploads/assets/articles/82136-new-gerber-baby-is-7-month-old-girl-named-grace/1421876680_grace-gerber-baby-zoom.jpg' WHERE age_group = 'Allentown'"
    cnnew.execute update_string
    response.write "UPDATED"
cnnew.close

else '**************************** couldn't open the user's database -- bail out!

       Response.write "<p>2. Open task failed on database: "
       Response.write cStr(fdsn)
       Response.write "<p>3. Drop current activity table NOT attempted"
       Response.write "<p>4. Create activity table NOT attempted"
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