<html>
<body>
<%@Page Language="C#"%>
<%@Import Namespace="System.Data" %>
<%@Import Namespace="System.Data.SqlClient" %>
<center><table border="1">
<tr>
<td><img src="captsm.gif"></td>
<td valign="middle" align="center" bgcolor="#aaaaaa"><font face="COMIC SANS MS"><font size="4">
<b>The Simple Table Listing Program Using a Data Grid<br>in c#<hr>Lists the Contents of <i>glmaster</i><br>Using a DataReader in SQL
</td>
</tr></table>

<div>SELECT command: <b><span id="outSelect" runat="server"></span></b></div>
<div id="outError" runat="server"> </div>
<div id="outResult" runat="server"></div>

<script language="C#" runat="server">

	void Page_Load(Object sender, EventArgs e)
	{
		// declare a string to hold the results as an HTML table
            
            string displayoutput = "<center><table>";
            int c=0;
            displayoutput=displayoutput+"<tr>";
            displayoutput=displayoutput+"<td align='center'>sequence</td>";
            displayoutput=displayoutput+"<td align='center'>uid</td>";
            displayoutput=displayoutput+"<td align='center'>time_stamp</td>";
            displayoutput=displayoutput+"<td align='center'>rating</td>";
            displayoutput=displayoutput+"<td align='center'>feedback</td>";


		string strConnect = "server=AUCKLAND;uid=gl1329;pwd=FRT46pzyF;database=gl1329";

		// specify the SELECT statement to extract the data

		string strSelect = "SELECT * FROM progress";
		outSelect.InnerText = strSelect;		// and display it

		try
		{
			// create a new Connection object using the connection string
			SqlConnection objConnect = new SqlConnection(strConnect);

			// open the connection to the database
			objConnect.Open();

			// create a new Command using the connection object and select statement
			SqlCommand objCommand = new SqlCommand(strSelect, objConnect);

			// declare a variable to hold a DataReader object
			SqlDataReader objDataReader;

			// execute the SQL statement against the command to fill the DataReader
			objDataReader = objCommand.ExecuteReader();

			// iterate through the records in the DataReader getting field values
			// the Read method returns False when there are no more records
			while (objDataReader.Read())
			{
			  displayoutput += "<tr><td>" + objDataReader["sequence"] + "</td><td>"
          				 + objDataReader["uid"] + "</td><td>"
				         + objDataReader["time_stamp"] + "</td><td>"
				         + objDataReader["rating"] + "</td><td>"
				         + objDataReader["feedback"] + "</td><td align='right'>";
			}

			// close the DataReader and Connection
			objDataReader.Close();
			objConnect.Close();
		}
		catch (Exception objError)
		{
			// display error details
			outError.InnerHtml = "<b>* Error while accessing data</b>.<br />"
							+ objError.Message + "<br />" + objError.Source;
			return;		//  and stop execution
		}
		DateTime t = DateTime.Now;
                displayoutput += "<P>Normal Termination "+t;

		outResult.InnerHtml = displayoutput;

	}
	
</script>
</body>
</html>
