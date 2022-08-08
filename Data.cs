using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGenerator
{
    class Data
    {
        public static DataTable SampleData()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add("Name");
            dt.Columns.Add("Salary");

            dt.Rows.Add("01", "Anis", "1000");
            dt.Rows.Add("02", "AK", "2000");
            dt.Rows.Add("03", "Mak", "3000");
            return dt;

        }
        
    }
}
