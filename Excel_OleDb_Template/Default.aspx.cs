using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data;
using System.Data.OleDb;
using Excel_OleDb_Template.DbFuntions;
namespace Excel_OleDb_Template
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string SQLText, ExcelTable, ExcelName;
            ExcelTable = "";
            ExcelName = DateTime.Now.ToString("yyMMddHHmmssffffff");
            using (FileSystemWatcher watcher = new FileSystemWatcher(DbSetting.ServerMapPath($"~/App_Data/")))
            {
                watcher.EnableRaisingEvents = true;
                watcher.IncludeSubdirectories = true;
                Response.Write("開始複製...<br>");
                File.Copy(DbSetting.ServerMapPath("~/App_Data/h.xlsx"), DbSetting.ServerMapPath($"~/App_Data/{ExcelName}.xlsx"));
                Response.Write("複製完成...<br>");
                watcher.Dispose();
            }

            
            try
            {
                using (OleDbConnection conn = new OleDbConnection(DbSetting.ServerMapPathExcel($"~/App_Data/{ExcelName}.xlsx")))
                {
                    conn.Open();
                    DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,new object[] { null, null, null, "TABLE" });
                    for (int i = 0; i < schemaTable.Columns.Count; i++)
                    {

                        if (schemaTable.Columns[i].ToString() == "TABLE_NAME")
                        {
                            ExcelTable = schemaTable.Rows[0][i].ToString();
                        }
                    }
                    
                    
                    using (OleDbCommand command = new OleDbCommand { Connection = conn, CommandText = $"insert into [{ExcelTable}] (管理,訂單編號,會員金額) values('123','456','789')" })
                    {
                        command.ExecuteNonQuery();
                        command.Cancel();
                        command.Dispose();
                    }
                    /*
                    //Response.Write(schemaTable.Columns.Count);
                    
                    //Response.Write(ExcelTable);
                    
                    SQLText = $"select * from [;database={DbSetting.ServerMapPath("/db/Database.accdb")};pwd=tn999kinggnik999nt].[commodity]";
                    //SQLText = $"select * from {ExcelTable}";
                    //SQLText = $"insert into [;database={DbSetting.ServerMapPath("/App_Data/test2.accdb")}].[B1] (B1,B2) select A4, A5 from [123$]";
                    using (OleDbCommand comm = new OleDbCommand { Connection = conn, CommandText = SQLText })
                    {
                        using (OleDbDataReader dr = comm.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    Response.Write(string.Format("{0}<br>", dr["commodityNumber"].ToString()));
                                }
                            }
                            dr.Close();
                        }
                        comm.Cancel();
                        comm.Dispose();
                    }
                    */
                    
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