
using System.Configuration;
using System.Data.OleDb;
using System;
using System.Diagnostics;
using System.Data;
using System.Threading.Tasks;
using System.Threading;
using System.Collections.Generic;

namespace FilterExcelData.AppCode
{
    internal class FilteredData
    {
        private static DataTable data = new DataTable();
        private static void LoadFromSingleFileSingleSheet()
        {
            var sw = new Stopwatch();
            sw.Start();

            var path = ConfigurationManager.AppSettings["path"];
            var file = ConfigurationManager.AppSettings["singleFileSingleSheet"];
            var sheet = ConfigurationManager.AppSettings["sheetName"];
            var columns = ConfigurationManager.AppSettings["columns"];
            var constraints = ConfigurationManager.AppSettings["constraints"];
            if (constraints != null && constraints != "")
            {
                constraints = "where " + string.Join(" and ", constraints.Split(','));
            }

            var dt = new DataTable();

            string conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + file + ";Extended Properties='Excel 12.0;HDR=YES';";

            Console.WriteLine("Fetching filtered data from single file single sheet..");
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter($"select {columns} from [{sheet}$]{constraints}", con);
                    oleAdpt.Fill(dt);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            sw.Stop();
            Program.DisplayTable(dt);
            Console.WriteLine("\n\nTotal Rows : " + dt.Rows.Count);
            Console.WriteLine("Time Elapshed In fetching the filtered data - \nTicks : " + sw.ElapsedTicks + "\nMilliseconds : " + sw.ElapsedMilliseconds);
            Console.ReadKey();
        }
        private static void LoadFromMultipleFilesUsingLoop()
        {
            var sw = new Stopwatch();
            sw.Start();

            var path = ConfigurationManager.AppSettings["path"];
            var files = ConfigurationManager.AppSettings["MultipleFiles"];
            var sheet = ConfigurationManager.AppSettings["sheetName"];
            var columns = ConfigurationManager.AppSettings["columns"];
            var constraints = ConfigurationManager.AppSettings["constraints"];
            if (constraints != null && constraints != "")
            {
                constraints = "where " + string.Join(" and ", constraints.Split(','));
            }
            
            Console.WriteLine("Fetching data from multiple files using foreach loop..");

            var dt = new DataTable();
            foreach(var file in files.Split(','))
            {
                string conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + file + ";Extended Properties='Excel 12.0;HDR=YES';";
                using (OleDbConnection con = new OleDbConnection(conn))
                {
                    try
                    {
                        OleDbDataAdapter oleAdpt = new OleDbDataAdapter($"select {columns} from [{sheet}$]{constraints}", con);
                        oleAdpt.Fill(dt);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                }

            }

            sw.Stop();
            Program.DisplayTable(dt);
            Console.WriteLine("\n\nTotal Rows : " + dt.Rows.Count);
            Console.WriteLine("Time Elapshed In fetching the filtered data - \nTicks : " + sw.ElapsedTicks + "\nMilliseconds : " + sw.ElapsedMilliseconds);
            Console.ReadKey();
        }
        private static async Task LoadFromFileAsync(string conn, string columns, string sheet, string constraints)
        {
            var dt = new DataTable();
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter($"select {columns} from [{sheet}$] {constraints}", con);
                    var task = Task.Run(() => { oleAdpt.Fill(dt); });
                    await task;
                    if(dt != null && data != null)
                        data.Merge(dt);
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
            var columns = ConfigurationManager.AppSettings["columns"];
            var constraints = ConfigurationManager.AppSettings["constraints"];
            if (constraints != null && constraints != "")
            {
                constraints = "where " + string.Join(" and ", constraints.Split(','));
            }

            Console.WriteLine("Fetching data from multiple files using Async/Await");

            var dt = new DataTable();
            var tasks = new List<Task>();
            foreach (var file in files.Split(','))
            {
                Console.WriteLine($"Fetching from {file}");
                string conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + file + ";Extended Properties='Excel 12.0;HDR=YES';";
                var task = LoadFromFileAsync(conn, columns, sheet, constraints);
                tasks.Add(task);
            }

            await Task.WhenAll(tasks);

            sw.Stop();

            Program.DisplayTable(data);
            Console.WriteLine("\n\nTotal Rows : " + data.Rows.Count);
            Console.WriteLine("Time Elapshed In fetching the filtered data - \nTicks : " + sw.ElapsedTicks + "\nMilliseconds : " + sw.ElapsedMilliseconds);
            Console.ReadKey();
        }
        private static void LoadFromMultipleFiles()
        {
            //using loop
            //LoadFromMultipleFilesUsingLoop();

            //using loop + async/await
            LoadFromMultipleFilesAsync().Wait();

        }

        internal static void MainCode()
        {
            Console.WriteLine("Application started. Loading the filtered data from multiple files using OleDB");

            //LoadFromSingleFileSingleSheet();

            LoadFromMultipleFiles();
        }
    }
}
