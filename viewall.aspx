
<%@Page Language="C#"%>
<%@Import Namespace="System.Data" %>
<%@Import Namespace="System.Data.SqlClient" %>
<html>
<body>
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
        displayoutput=displayoutput+"<td align='center'>activity_id</td>";
        displayoutput=displayoutput+"<td align='center'>a_name</td>";
        displayoutput=displayoutput+"<td align='center'>age_group</td>";
        displayoutput=displayoutput+"<td align='center'>theme</td>";


        string strConnect = "server=AUCKLAND;uid=gl1329;pwd=FRT46pzyF;database=gl1329";

        // specify the SELECT statement to extract the data

        string strSelect = "SELECT * FROM activity";
        outSelect.InnerText = strSelect;        // and display it

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
            int r=0;
            while (objDataReader.Read())
            {
                displayoutput += "<tr><td>" + objDataReader["activity_id"] + "</td><td>"
                             + objDataReader["a_name"] + "</td><td>"
                           + objDataReader["age_group"] + "</td><td>"
                           + objDataReader["theme"] + "</td><td>";
                r=r+1;
            }
            outResult.InnerHtml = "ROWS:" + r.ToString();
            // close the DataReader and Connection
            objDataReader.Close();
            objConnect.Close();
        }
        catch (Exception objError)
        {
            // display error details
            outError.InnerHtml = "<b>* Error while accessing data</b>.<br />"
                            + objError.Message + "<br />" + objError.Source;
            return;     //  and stop execution
        }
        DateTime t = DateTime.Now;
        displayoutput += "<P>Normal Termination "+t;

        outResult.InnerHtml = displayoutput;

    }

</script>
</body>
</html>
