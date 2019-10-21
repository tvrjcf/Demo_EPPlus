using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Demo_EPPlus
{
    public class Test
    {
        public string createFile(string fileName)
        {
            try
            {
                Console.WriteLine("start...");
                fileName = string.IsNullOrEmpty(fileName) ? @"C:\sample1.xlsx" : fileName;
                FileInfo newFile = new FileInfo(fileName);
                ExcelPackage excel = new ExcelPackage(newFile);
                //Add the Content sheet
                var ws = excel.Workbook.Worksheets["Content"];
                if (ws == null)
                {
                    ws = excel.Workbook.Worksheets.Add("Content");
                    ws.Cells["B2"].Value = "Demo_EPPlus";
                }
                else
                    ws.Cells["B2"].Value = "Demo_EPPlus 2";
                excel.Save();
                //var fullName = Path.GetFullPath(fileName);
                //Console.WriteLine(fileName);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.GetBaseException().Message);
            }
            return Path.GetFullPath(fileName);
        }
    }
}
