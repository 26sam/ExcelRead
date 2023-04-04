
using System.Configuration;
using System.Data.OleDb;
using System.Data;
using System.Diagnostics;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.IO;
using LinqToExcel;

namespace FilterExcelData.AppCode
{
    internal class FilterUsingLINQ
    {
        private static List<DataRow> dataRows = new List<DataRow>();
        private static void PrintRecord(IQueryable<Row> record, string columns)
        {
            foreach (Row row in record)
            {
                foreach (var column in columns.Split(','))
                {
                    Console.Write($"{row[column], -35}");
                }
                Console.WriteLine('\n');
            } 
        }
        private static void PrintData(string columns)
        {
            foreach (DataRow row in dataRows)
            {
                try
                {
                    foreach(var column in columns.Split(','))
                    {
                        Console.Write($"{row[column], -35}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                Console.WriteLine('\n');
            }

        }
        private static void FilterData(DataTable dataTable)
        {
            foreach (DataRow row in dataTable.Rows)
            {
                if (row["first_name"].ToString() == "Von")
                {
                    dataRows.Add(row);
                }
            }
        }
        private static async Task LoadFromFileAsync(string connectionAddress, string columns, string sheet, string constraints)
        {
            DataTable dataTable = new DataTable();
            using (OleDbConnection oleConnection = new OleDbConnection(connectionAddress))
            {
                try
                {
                    DataSet ds = new DataSet();
                    OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$]", oleConnection);
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter(cmd);
                    var task = Task.Run(() => { oleAdpt.Fill(ds); });
                    await task;
                    FilterData(ds.Tables[sheet]);
                }
                catch(Exception ex)
                {
                    Console.WriteLine("Exception: " + ex.Message);
                }
                /*try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter($"select {columns} from [{sheet}$] {constraints}", oleConnection);
                    var task = Task.Run(() => { oleAdpt.Fill(dataTable); });
                    await task;
                    FilterData(dataTable);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception: " + ex.Message);
                }*/
            }
            PrintData(columns);
        }
        private static void LoadFromSingleFileMultipleSheets()
        {
            var stopWatch = new Stopwatch();
            stopWatch.Start();

            var filePath = ConfigurationManager.AppSettings["singleFile"];
            //var sheets = ConfigurationManager.AppSettings["multipleSheets"];
            var displayingColumns = ConfigurationManager.AppSettings["columns"];

            ExcelQueryFactory factory = new ExcelQueryFactory(filePath);
            var excelSheets = factory.GetWorksheetNames().ToList();

            IQueryable<Row> record = null;

            for(int i=0;i < excelSheets.Count;i++)
            {
                Console.WriteLine("Sheet"+(i+1));
                record = factory.Worksheet(excelSheets[i]).Select(row => row);
                record = record.Where(row => row["first_name"] == "Von");
                PrintRecord(record, displayingColumns);
            }
            stopWatch.Stop();
            Console.WriteLine("Time elapsed in retrieving the filtered data - \nTicks : " + stopWatch.ElapsedTicks + "\nMilliseconds : " + stopWatch.ElapsedMilliseconds);
            Console.ReadKey();

        }
        private static void LoadFromSingleFileSingleSheet()
        {
            var stopWatch = new Stopwatch();
            stopWatch.Start();

            //var path = ConfigurationManager.AppSettings["path"];
            var file = ConfigurationManager.AppSettings["singleFile"];
            var sheet = ConfigurationManager.AppSettings["sheetName"];
            var displayingColumns = ConfigurationManager.AppSettings["columns"];
            var constraints = ConfigurationManager.AppSettings["constraints"];

            if (constraints != null && constraints != "")
            {
                constraints = "where " + string.Join(" and ", constraints.Split(','));
            }

            var dataTable = new DataTable();

            string connectionAddress = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file + ";Extended Properties='Excel 12.0;HDR=YES';";

            Console.WriteLine("Data from single file single sheet......");
            using (OleDbConnection oleConnection = new OleDbConnection(connectionAddress))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter($"select * from [{sheet}$]", oleConnection);
                    oleAdpt.Fill(dataTable);
                    FilterData(dataTable);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            stopWatch.Stop();

            PrintData(displayingColumns);
            Console.WriteLine("Time elapsed in fetching the filtered data - \nTicks : " + stopWatch.ElapsedTicks + "\nMilliseconds : " + stopWatch.ElapsedMilliseconds);
            Console.ReadKey();
        }
        private static async Task LoadFromMultipleFiles()
        {
            var stopWatch = new Stopwatch();
            stopWatch.Start();

            var path = ConfigurationManager.AppSettings["path"];
            var files = ConfigurationManager.AppSettings["multipleFiles"];
            var sheet = ConfigurationManager.AppSettings["sheetName"];
            var displayingColumns = ConfigurationManager.AppSettings["columns"];
            var constraints = ConfigurationManager.AppSettings["constraints"];

            if (constraints != null && constraints != string.Empty)
            {
                constraints = "where " + string.Join(" and ", constraints.Split(','));
            }

            Console.WriteLine("Data from multiple files......");

            var tasks = new List<Task>();
            int count = 1;
            foreach (var file in files.Split(','))
            {
                Console.WriteLine($"Read from File {count++}");
                string connectionAddress = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + file + ";Extended Properties='Excel 12.0;HDR=YES';";

                var task = LoadFromFileAsync(connectionAddress, displayingColumns, sheet, constraints);
                tasks.Add(task);
            }

            await Task.WhenAll(tasks);

            stopWatch.Stop();

            PrintData(displayingColumns);
            Console.WriteLine("Total time in fetching the filtered data - \nTicks : " + stopWatch.ElapsedTicks + "\nMilliseconds : " + stopWatch.ElapsedMilliseconds);
            Console.ReadKey();

        }
        public static void MainCode()
        {
            //LoadFromMultipleFiles().Wait();
            //LoadFromSingleFileSingleSheet();
            LoadFromSingleFileMultipleSheets();
        }
    }
}
