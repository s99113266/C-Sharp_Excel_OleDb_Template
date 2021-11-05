using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;
using Excel_OleDb_Template.DbFuntions;
namespace Excel_OleDb_Template
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string SQLText, ExcelTable;
            ExcelTable = "";
            try
            {
                using (OleDbConnection conn = new OleDbConnection(DbSetting.AccdbConnStr1))
                {
                    conn.Open();
                    DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,new object[] { null, null, null, "TABLE" });


                    //Response.Write(schemaTable.Columns.Count);
                    
                    for (int i = 0; i < schemaTable.Columns.Count; i++)
                    {

                        if (schemaTable.Columns[i].ToString() == "TABLE_NAME")
                        {
                            //Response.Write(schemaTable.Columns[i].ToString() + " : " + schemaTable.Rows[0][i].ToString() + "<br>");
                            ExcelTable = schemaTable.Rows[0][i].ToString();
                        }
                        //Response.Write(schemaTable.Columns[i].ToString() + "<br>");
                    }
                    //Response.Write(ExcelTable);
                    
                    SQLText = $"select * from [;database={DbSetting.ServerMapPath("/App_Data/test2.accdb")}].[B1]";
                    SQLText = $"select * from {ExcelTable}";
                    //SQLText = $"insert into [;database={DbSetting.ServerMapPath("/App_Data/test2.accdb")}].[B1] (B1,B2) select A4, A5 from [123$]";
                    using (OleDbCommand comm = new OleDbCommand { Connection = conn, CommandText = SQLText })
                    {
                        using (OleDbDataReader dr = comm.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                //Response.Write(dr.FieldCount + "<br>");
                                while (dr.Read())
                                {
                                    Response.Write(string.Format("{0}<br>", dr["B2"].ToString()));
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