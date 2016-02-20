<%@Page Language="C#" Debug="true"%>
<%@Import Namespace="System.Data" %>
<%@Import Namespace="System.Data.SqlClient" %>
<%@Import Namespace="System.Text.RegularExpressions" %>
<html>
<script language="C#" runat="server">
    void Page_Load(Object sender, EventArgs e)
    {
    }
    void test()
    {
        string errortext;
        try
        {
            int numa;
            int oksw;
            string adtemp;
            string adtemp2;
            string strUpdate;

            string strInsert;
            string strConnect;
            int c;
            int uid;
            int flag;
            string feedback;
            strInsert = "INSERT INTO progress (sequence, uid, time_stamp, rating, feedback) VALUES (";
            strConnect = "server=AUCKLAND;database=gl1329;uid=gl1329;pwd=FRT46pzyF";
            SqlConnection myConnection = new SqlConnection(strConnect);
            SqlCommand myCommand = new SqlCommand(strInsert, myConnection);
            //
            //*** Server Side Regular Expression edit of account description text field
            //
            c = 0;
            uid = 1000;
            feedback = "TestData";
            myConnection.Open();
            for (int i = 0; i < 100; i++)
            {
                c = c + 1;
                uid = uid + 1;
                flag = 0;
                feedback = feedback + i.ToString();
                strInsert = strInsert + "'" + c.ToString() + "'" + "," + "'" + uid.ToString() + "'" + "," + "'" + DateTime.Now + "'" + "," + "'" + flag.ToString() + "'" + "," + "'" + feedback + "'" + ")";
                flag = flag - 1;
                numa = myCommand.ExecuteNonQuery();
            }
            results.InnerHtml = "";
            error_out.InnerHtml = "";
            myConnection.Close();
            //if (oksw==0)
            //{
            //    myCommand.Parameters.AddWithValue( "@adv"   , adtemp );
            //    myCommand.Parameters.AddWithValue( "@mjrv"  , mjr.Text );
            //    myCommand.Parameters.AddWithValue( "@mnrv"  , mnr.Text );
            //    myCommand.Parameters.AddWithValue( "@s1v"   , s1.Text );
            //    myCommand.Parameters.AddWithValue( "@s2v"   , s2.Text );
            //    myConnection.Open();
            //    numa = myCommand.ExecuteNonQuery();
            //    results.InnerHtml=numa.ToString() + " record inserted OK.";
            //    myConnection.Close();
            //    ClearForm();
            //}
        }
        catch (Exception exc)
        {
            errortext = exc.Message.ToUpper();
            if (errortext.IndexOf("DUPLICATE") > 0)
                error_out.InnerHtml = "The Account you are trying to add <b>ALREADY EXISTS</b> in the databse and cannot be duplicated.<br>Check your data and try again.";
            else
          if (errortext.IndexOf("DATA TYPE") > 0)
                error_out.InnerHtml = "Invalid Data Entered. The data was NOT added to the database. Correct the data problem above and try again";
            else
                error_out.InnerHtml = "<b>* Error while Inserting</b>.<br />" + exc.Message + "<br />" + exc.Source;
        }

    }

</script>
    <font color="#00dd00"><DIV id="results" Runat="server"></DIV></font>
<font color="#ff0000"><DIV id="error_out" Runat="server"></DIV></font>
</body>
</html>