using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.OleDb;
using Excel_OleDb_Template.DbFuntions;
namespace Excel_OleDb_Template
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string SQLText;
            SQLText = $"select * from [;database={DbSetting.ServerMapPath("/App_Data/test2.accdb")}].[B1]";
            //SQLText = $"insert into [;database={DbSetting.ServerMapPath("/App_Data/test2.accdb")}].[B1] (B1,B2) select A4, A5 from [123$]";
            try
            {
                using (OleDbConnection conn = new OleDbConnection(DbSetting.ExcelConnStr1))
                {
                    conn.Open();
                    //OleDbSchemaGuid schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,new object[] { null, null, null, "TABLE" });
                    using (OleDbCommand comm = new OleDbCommand { Connection = conn, CommandText = SQLText })
                    {
                        using (OleDbDataReader dr = comm.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                Response.Write(dr.FieldCount + "<br>");
                                while (dr.Read())
                                {
                                    Response.Write(dr["B1"].ToString());
                                }
                            }
                            dr.Close();
                        }
                        comm.Cancel();
                        comm.Dispose();
                    }
                    conn.Close();
                    conn.Dispose();
                }
            }
            catch (Exception err)
            {
                Response.Write(err.Message);
            }
        }
    }
}