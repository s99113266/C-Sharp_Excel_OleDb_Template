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
        public static string ServerMapPathExcel(string e)
        {
            return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={Context.Server.MapPath(e)};Extended Properties='Excel 12.0;HDR=Yes;'";
        }
    }
}