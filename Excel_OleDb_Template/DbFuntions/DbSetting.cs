using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
namespace Excel_OleDb_Template.DbFuntions
{
    public class DbSetting
    {
        public static HttpContext Context = HttpContext.Current;

        public static string ServerMapPath(string e)
        {
            return Context.Server.MapPath(e);
        }
        public static string ExcelConnStr1 = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={ServerMapPath("~/App_Data/test1.xlsx")};Extended Properties='Excel 12.0;HDR=Yes;IMEX=2;'";
    }
}