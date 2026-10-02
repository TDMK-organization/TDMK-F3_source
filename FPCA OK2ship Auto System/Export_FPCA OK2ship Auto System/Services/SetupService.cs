using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace Export_FPCA_OK2ship_Auto_System.Services
{
    public class SetupService
    {
        public static string[] ScanWorkSheet(string location)
        {
            location = location.Trim();
            if (!File.Exists(location))
            {
                throw new Exception("File not found");
            }

            List<string> list = new List<string>();
            ExportProcess _process = new ExportProcess();
            using (ExcelPackage package = _process.OpenFileExcel(location))
            {
                ExcelWorksheets  wss = package.Workbook.Worksheets;
                foreach (ExcelWorksheet ws in wss)
                {
                    list.Add(ws.Name);
                }
            }
            
            return list.ToArray();
        }
    }
}