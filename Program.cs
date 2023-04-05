using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using FilterExcelData.AppCode;
using System.Data;

namespace FilterExcelData
{
    internal class Program
    {
        internal static void DisplayTable(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    Console.Write($"{item,-15}");
                }
                Console.WriteLine();
            }
        }

        internal static int DisplayList(List<DataRow> ls)
        {
            int count = 0;
            foreach (var row in ls)
            {
                count++;
                try
                {
                    Console.Write($"{row["ItemType"],-15}{row["SalesChannel"],-15}{row["UnitsSold"],-15}{row["TotalProfit"],-15}\n");
                }
                catch { }
            }
            return count;
        }

        static void Main(string[] args)
        {
            FilteredData.MainCode();
            Console.Clear();
            FilterUsingLINQ.MainCode();
        }
    }
}
