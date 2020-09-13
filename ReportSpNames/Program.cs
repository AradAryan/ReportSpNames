using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;

namespace ReportSpNames
{
    internal static class Program
    {
        static void Main(string[] args)
        {
            var lines = File.ReadAllLines(@"C:\Users\javad\Desktop\DepositEntity.Context.cs");
            List<string> list = new List<string>();
            string name = "", param = "";
            foreach (var line in lines)
            {
                if (line.Contains("public virtual ObjectResult<"))
                {
                    param = line.Split('(')[1].Split(')')[0];
                    param = param.Replace("Nullable<", "").Replace(">", "");
                }
                if (line.Contains("ExecuteFunction"))
                {
                    name = line.Split('(')[3].Split('\"')[1];
                    if (!string.IsNullOrEmpty(param))
                        list.Add($"{name},{param}");
                    else
                        list.Add($"{name}");
                }
            }
            foreach (var item in list)
            {
                Console.WriteLine(item);
            }

            var fileName = string.Empty;
            var path = @"C:\Users\javad\Desktop\temp.xlsx";
            using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                using (var excel = new ExcelPackage(fileStream))
                {
                    fileStream.Close();
                    fileStream.Dispose();
                    #region [STYLE]
                    var workbook = excel.Workbook;
                    var worksheet = excel.Workbook.Worksheets[1];
                    #endregion

                    var RowIndex = 7;
                    if (list.Count() > 0)
                    {
                        foreach (var item in list)
                        {
                            var temp = item.Split(',');
                            worksheet.Cells[RowIndex, 3].Value = RowIndex - 6;
                            worksheet.Cells[RowIndex, 6].Value = temp[0];
                            RowIndex++;
                        }
                    }

                    File.WriteAllBytes(@"C:\Users\javad\Desktop\New folder\test.xlsx", excel.GetAsByteArray());
                }
            }

            Console.ReadLine();
        }
    }
}