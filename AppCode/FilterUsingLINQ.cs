
using System.Configuration;
using System.Data.OleDb;
using System.Data;
using System.Diagnostics;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace FilterExcelData.AppCode
{
    internal class FilterUsingLINQ
    {

        private static List<DataRow> FilterData(DataTable dt)
        {
            var data = new List<DataRow>();
            foreach (DataRow row in dt.Rows)
            {
                if (row["ItemType"].ToString() == "Snacks" && int.Parse(row["UnitsSold"].ToString()) > 9000)
                {
                    data.Add(row);
                }
            }
            return data;
        }
        private static void LoadFromSingleFileSingleSheet()
        {
            var sw = new Stopwatch();
            sw.Start();

            var path = ConfigurationManager.AppSettings["path"];
            var file = ConfigurationManager.AppSettings["singleFileSingleSheet"];
            var sheet = ConfigurationManager.AppSettings["sheetName"];

            var dt = new DataTable();
            string conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + file + ";Extended Properties='Excel 12.0;HDR=YES';";

            Console.WriteLine("Fetching data from single file single sheet..");

            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter($"select * from [{sheet}$]", con);
                    oleAdpt.Fill(dt);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }

            Console.WriteLine("Filtering the data...");

            int count = 0;
            var data = FilterData(dt);

            sw.Stop();
            Program.DisplayList(data);
            Console.WriteLine("\n\nTotal Rows : " + data.Count());
            Console.WriteLine("Time Elapshed In fetching and filtering the data - \nTicks : " + sw.ElapsedTicks + "\nMilliseconds : " + sw.ElapsedMilliseconds);
            Console.ReadKey();
        }
        
        private static List<DataRow> dataRows = new List<DataRow>();
        private static void FilterData(DataTable dt, string file)
        {
            Console.WriteLine($"Filtering Data of file : {file}");
            foreach (DataRow row in dt.Rows)
            {
                if (row["ItemType"].ToString() == "Snacks" && int.Parse(row["UnitsSold"].ToString()) > 9000)
                {
                    dataRows.Add(row);
                    //result.Rows.Add(row);
                }
            }
        }
        private static async Task LoadFromFileAsync(string conn, string sheet, string file)
        {
            var dt = new DataTable();
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter($"select * from [{sheet}$]", con);
                    var task = Task.Run(() => { oleAdpt.Fill(dt); });
                    await task;
                    task = Task.Run(() => { FilterData(dt, file); });
                    await task;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }
        private static async Task LoadFromMultipleFilesAsync()
        {
            var sw = new Stopwatch();
            sw.Start();

            var path = ConfigurationManager.AppSettings["path"];
            var files = ConfigurationManager.AppSettings["MultipleFiles"];
            var sheet = ConfigurationManager.AppSettings["sheetName"];
            
            Console.WriteLine("Fetching data from multiple files using Async/Await");

            var tasks = new List<Task>();
            foreach (var file in files.Split(','))
            {
                Console.WriteLine($"Fetching from {file}");
                string conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + file + ";Extended Properties='Excel 12.0;HDR=YES';";
                var task = LoadFromFileAsync(conn, sheet,file);
                tasks.Add(task);
            }

            await Task.WhenAll(tasks);
            sw.Stop();

            int count = Program.DisplayList(dataRows);
            Console.WriteLine("\n\nTotal Rows : " + dataRows.Count + "  count : " + count);
            Console.WriteLine("Time Elapshed In fetching and filtering data - \nTicks : " + sw.ElapsedTicks + "\nMilliseconds : " + sw.ElapsedMilliseconds);
            Console.ReadKey();

        }

        public static void MainCode()
        {
            Console.WriteLine("Application Started : First fetching the data from excel files using OleDB and then filtering the data using loop\n");

            //LoadFromSingleFileSingleSheet();

            LoadFromMultipleFilesAsync().Wait();
        }
    }
}
