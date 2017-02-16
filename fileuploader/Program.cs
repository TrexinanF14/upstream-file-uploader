using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using LumenWorks.Framework.IO.Csv;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;

namespace fileuploader
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Welcome to the Current file uploader!");
            Console.WriteLine("This uploader expects an excel or csv-formatted file, with the first row containing column headers matching the names of the fields for your channel.");

            string filepath = GetArgAfter(args, "--filename");

            while (!File.Exists(filepath))
            {
                if (filepath != null)
                {
                    Console.WriteLine("The file specified doesn't exist, please try again:");
                }
                Console.WriteLine("Please enter the name of the file to upload: ");
                filepath = Console.ReadLine();
            }
            Console.WriteLine("Reading rows from file...");
            var rows = ReadFileRows(filepath);
            if (rows == null)
            {
                return;
            }
            Console.WriteLine(rows.Count + " rows found in file (not including the header row)");

            Uri url = null;
            try
            {
                string urlstr = GetArgAfter(args, "--webhook");
                if (urlstr != null)
                {
                    url = new Uri(urlstr);
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Invalid webhook url parameter.");
            }

            while (url == null)
            {
                try
                {
                    Console.WriteLine("Enter the webhook url for the channel you want to upload to:");
                    url = new Uri(Console.ReadLine());
                }
                catch (Exception)
                {
                    Console.WriteLine("Invalid webhook url. Make sure the full url is getting copied.");
                }
            }

            double pause;
            if (!double.TryParse(GetArgAfter(args, "--pause"), out pause))
            {
                Console.WriteLine("Enter the pause time in seconds between record uploads (0 or just hit enter for no pause)");
                var output = Console.ReadLine();
                if (!double.TryParse(output, out pause))
                {
                    pause = 0;
                }
            }

            Console.WriteLine("Starting to upload rows to Current....");

            UploadRows(rows, url, pause).Wait();
            Console.WriteLine("Finished uploading.");
        }

        static string GetArgAfter(string[] args, string token)
        {
            if (args.Contains(token))
            {
                int idx = Array.FindIndex(args, arg => arg == token);
                if (idx + 1 >= args.Length)
                {
                    return null;
                }
                return args[idx + 1];
            }
            return null;
        }

        static List<Dictionary<string, object>> ReadFileRows(string filepath)
        {
            if (filepath.ToLower().EndsWith(".csv"))
            {
                return ReadCSVFile(filepath);
            }
            else if (filepath.ToLower().EndsWith(".xls") || filepath.ToLower().EndsWith(".xlsx"))
            {
                return ReadExcelFile(filepath);
            }
            else
            {
                Console.WriteLine("You passed a file with an invalid file extension. The valid file types are .csv, .xls, and .xlsx");
                return null;
            }
        }

        static List<Dictionary<string, object>> ReadExcelFile(string filepath)
        {
            using (FileStream stream = File.Open(filepath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader excelReader;
                if (filepath.ToLower().EndsWith(".xls"))
                {
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else
                {
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                var ds = excelReader.AsDataSet();
                var headers = ds.Tables[0].Rows[0].ItemArray
                    .Select((x, index) => new KeyValuePair<int, string>(index, x?.ToString()))
                    .Where(x => !string.IsNullOrEmpty(x.Value))
                    .ToDictionary(x => x.Key, x => x.Value);

                var rows = new List<Dictionary<string, object>>();
                for (int rowindex = 1; rowindex < ds.Tables[0].Rows.Count; rowindex++)
                {
                    var row = new Dictionary<string, object>();
                    foreach (var kvp in headers)
                    {
                        var rowval = ds.Tables[0].Rows[rowindex][kvp.Key]?.ToString();
                        double val;
                        if (double.TryParse(rowval, out val))
                        {
                            row.Add(kvp.Value, val);
                        }
                        else
                        {
                            row.Add(kvp.Value, rowval);
                        }
                    }
                    rows.Add(row);
                }
                return rows;
            }
        }

        static List<Dictionary<string, object>> ReadCSVFile(string filepath)
        {
            var rows = new List<Dictionary<string, object>>();
            using (CsvReader csv = new CsvReader(new StreamReader(filepath), true))
            {
                string[] headers = csv.GetFieldHeaders();
                while (csv.ReadNextRecord())
                {
                    var row = new Dictionary<string, object>();
                    for (int i = 0; i < headers.Length; i++)
                    {
                        double val;
                        if (double.TryParse(csv[i], out val))
                        {
                            row.Add(headers[i], val);
                        }
                        else
                        {
                            row.Add(headers[i], csv[i]);
                        }
                    }
                    rows.Add(row);
                }
            }
            return rows;
        }

        static async Task UploadRows(List<Dictionary<string, object>> rows, Uri url, double pause)
        {
            using (var client = new HttpClient())
            {
                client.BaseAddress = url;
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                if (pause == 0)
                {
                    await client.PostAsync(url.PathAndQuery, new StringContent(JsonConvert.SerializeObject(rows)));
                }
                else
                {
                    int count = 1;
                    foreach (var row in rows)
                    {
                        await client.PostAsync(url.PathAndQuery, new StringContent(JsonConvert.SerializeObject(row)));
                        Console.WriteLine("Sent row " + count);
                        count++;
                        System.Threading.Thread.Sleep(Convert.ToInt32((double)pause * 1000));
                    }
                }
            }
        }
    }

    public class FileDescriptor
    {
        public List<string> Fields;

        public List<Dictionary<string, string>> Rows;
    }
}
