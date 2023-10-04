using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Configuration;
using BIG;

namespace OrderSum
{
    public class Events
    {
        internal BIG.Application bigApp;

        public static Configuration appConfig = ConfigurationManager.OpenExeConfiguration(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OrderSum.dll"));

        public static string connectionString = appConfig.ConnectionStrings.ConnectionStrings["vbConnection"].ConnectionString;

        public Events(BIG.Application bigApplication)
        {
            this.bigApp = bigApplication;

            this.bigApp.SessionLoggedOn += SessionLoggedOn;

            this.bigApp.ButtonClick += ButtonClick;
        }

        private void SessionLoggedOn()
        {
            // Logged on user   
            BIG.Ts_ReadOnlyRowStringValue ts;
            string username = bigApp.User.GetStringValue((int)C_User.C_User_UserName, out ts);

            Log("");
            Log("Logged on! - " + DateTime.Now.ToString() + " - " + username);
        }

        private void ButtonClick(BIG.Button button, ref bool SkipRecording)
        {
            try
            {
                if (button.Caption.ToLower() == "summer")
                {
                    BIG.PageElement pageElement = button.PageElement;

                    BIG.Document bigDocument = pageElement.Document;

                    string commandText = @"
                        UPDATE FreeInf1  
                           SET 
                            Val2 = (SELECT SUM(NoInvo + NoFin) FROM OrdLn WHERE ProdNo IN (SELECT ProdNo FROM Prod WHERE Gr2 = 1 AND ProdNo = FreeInf1.ProdNo)),
                            Val3 = (SELECT SUM(NoFin) FROM OrdLn WHERE ProdNo IN (SELECT ProdNo FROM Prod WHERE Gr2 = 1 AND ProdNo = FreeInf1.ProdNo))";

                    using (SqlConnection sqlConnection = new SqlConnection(connectionString))
                    {
                        using (SqlCommand sqlCommand = new SqlCommand(commandText))
                        {
                            sqlConnection.Open();
                            sqlCommand.Connection = sqlConnection;
                            sqlCommand.ExecuteNonQuery();
                        }
                    }

                    bigDocument.Refresh();
                }
            }
            catch (Exception ex)
            {
                Log("Exception error: " +  ex.Message);
            }
        }

        public static void Log(string message)
        {
            StreamWriter sw = null;

            try
            {
                sw = new StreamWriter("OrderSumLog.txt", true);
                sw.WriteLine(DateTime.Now.ToString() + ": " + message);
                sw.Flush();
                sw.Close();
            }
            catch
            {
            }
        }
    }
}

