using OfficeOpenXml;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace OK2SHIP_SMT.Services
{
    public class GAPConnectorService
    {
        private DBContext _dbContext = new DBContext();
        public void Export(ExcelWorksheet workSheet, DataTable dataTable, DataTable spec)
        {

            int SpecNum = int.Parse(spec.Rows[0]["Count_Sample"].ToString());
            string[] healder = new[] { "IO PIN", "GROUNDING PIN_LEFT", "GROUNDING PIN_RIGHT", "Sample 1", "Average", "Min", "Max" };
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(workSheet, healder);

            Dictionary<string, string> _heal = new Dictionary<string, string>();
            foreach (string item in new[] { "IO PIN", "GROUNDING PIN_LEFT", "GROUNDING PIN_RIGHT" })
            {
                if (dic.TryGetValue("Sample 1", out string addressS))
                {
                    string addRes = "";
                    int min = int.MaxValue;
                    foreach (string jtem in addressS.Split('-'))
                    {
                        int distance = ExportProcess.DistanceRow(dic[item], jtem);
                        if (distance < min && distance > 0)
                        {
                            addRes = jtem;
                            min = distance;
                        }
                    }
                    _heal.Add(dic[item], addRes);
                }
            }
            foreach (string item in new[] { "IO PIN", "GROUNDING PIN_LEFT", "GROUNDING PIN_RIGHT" })
            {
                DataTable originalTable = new DataTable();
                switch (item)
                {
                    case "IO PIN":
                        originalTable = dataTable.AsEnumerable()
                                                      .Where(row => row.Field<string>("region") == "GAP DOC")
                                                    .CopyToDataTable();
                        break;
                    case "GROUNDING PIN_LEFT":
                        originalTable = dataTable.AsEnumerable()
                                                  .Where(row => row.Field<string>("region") == "GAP TRU")
                                                  .OrderBy(row =>
                                                  {
                                                      string sampleValue = row.Field<string>("Sample");
                                                      // Kiểm tra và chuyển đổi. Nếu chuyển đổi thất bại hoặc null/empty, coi là 0 (hoặc giá trị thấp nhất)
                                                      if (int.TryParse(sampleValue, out int result))
                                                      {
                                                          return result;
                                                      }
                                                      return int.MinValue;
                                                  })
                                                  .Take(5)
                                                .CopyToDataTable();
                        break;
                    case "GROUNDING PIN_RIGHT":
                        originalTable = dataTable.AsEnumerable()
                                                  .Where(row => row.Field<string>("region") == "GAP TRU")
                                                     .OrderBy(row =>
                                                     {
                                                         string sampleValue = row.Field<string>("Sample");
                                                         // Kiểm tra và chuyển đổi. Nếu chuyển đổi thất bại hoặc null/empty, coi là 0 (hoặc giá trị thấp nhất)
                                                         if (int.TryParse(sampleValue, out int result))
                                                         {
                                                             return result;
                                                         }
                                                         return int.MinValue;
                                                     })
                                                     .Skip(5)
                                                  .Take(5)
                                                .CopyToDataTable();
                        break;

                }
                bool prime = false;
                string address = _heal[dic[item]];
                for (int i = 0; i < SpecNum; i++)
                {
                    string addressP = ExportProcess.AddColumn(address, i);
                    // Kiểm tra và ghi product ID
                    if (prime || workSheet.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressP, -1), -1)].Text.Contains("Flex"))
                    {
                        prime = true;
                        if (originalTable.Columns.Contains("ProductID"))
                        {
                            workSheet.Cells[ExportProcess.AddRow(addressP, -1)].Value = originalTable.Rows[i]["ProductID"];
                        }
                    }
                    else
                    {
                        addressP = ExportProcess.AddRow(addressP, -1);
                    }

                    addressP = ExportProcess.AddRow(addressP, 1);
                    bool imageMode = !originalTable.Columns["Image"].GetType().ToString().Contains("byte");
                    if (imageMode)
                    {
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressP], TDMK_ImageConverter.ImageToByteArray((Image)originalTable.Rows[i]["Image"], ImageFormat.Jpeg), $"{Guid.NewGuid()}");

                    }
                    else
                    {
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressP], (byte[])originalTable.Rows[i]["Image"], $"{Guid.NewGuid()}");
                    }
                    addressP = ExportProcess.AddRow(addressP, 1);
                    if (imageMode)
                    {
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressP], TDMK_ImageConverter.ImageToByteArray((Image)originalTable.Rows[i]["Image1"], ImageFormat.Jpeg), $"{Guid.NewGuid()}");

                    }
                    else
                    {
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressP], (byte[])originalTable.Rows[i]["Image1"], $"{Guid.NewGuid()}");
                    }
                    string[] data1 = originalTable.Rows[i]["Data"].ToString().Split('/')[0].Trim().TrimEnd(';').Split(';');
                    try
                    {
                        for (int j = 0; j < data1.Count(); j++)
                        {
                            addressP = ExportProcess.AddRow(addressP, 1);
                            workSheet.Cells[addressP].Value = float.Parse(data1[j]);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Lỗi khi ghi data {ex.Message}");
                    }
                    addressP = ExportProcess.AddRow(addressP, 1);
                    if (imageMode)
                    {
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressP], TDMK_ImageConverter.ImageToByteArray((Image)originalTable.Rows[i]["Image2"], ImageFormat.Jpeg), $"{Guid.NewGuid()}");

                    }
                    else
                    {
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressP], (byte[])originalTable.Rows[i]["Image2"], $"{Guid.NewGuid()}");
                    }
                    string[] data2 = originalTable.Rows[i]["Data"].ToString().Split('/')[1].Trim().TrimEnd(';').Split(';');
                    try
                    {
                        for (int j = 0; j < data2.Count(); j++)
                        {
                            addressP = ExportProcess.AddRow(addressP, 1);
                            workSheet.Cells[addressP].Value = float.Parse(data2[j]);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Lỗi khi ghi data {ex.Message}");
                    }

                    addressP = ExportProcess.AddRow(addressP, 1);
                    workSheet.Cells[addressP].Value = "OK";

                }

            }
            foreach (string key in new[] { "Average", "Min", "Max" })
            {
                foreach (string adress in dic[key].Split('-'))
                {
                    for (int i = 0; i < 4; i++)
                    {
                        string address = ExportProcess.AddRow(adress, i + 1);
                        workSheet.Cells[address].Value = 0.1;
                        switch (key)
                        {
                            case "Average":
                                workSheet.Cells[address].FormulaR1C1 = $"=AVERAGE(RC[-5]:RC[-1])";
                                break;
                            case "Min":
                                workSheet.Cells[address].FormulaR1C1 = $"=min(RC[-6]:RC[-2])";
                                break;
                            case "Max":
                                workSheet.Cells[address].FormulaR1C1 = $"=max(RC[-7]:RC[-3])";
                                break;
                            default:
                                Debugger.Break();
                                break;
                        }

                    }
                }
            }


        }
        public static DataTable getConstructor()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("ItemCode", typeof(string));
            dt.Columns.Add("LotNo", typeof(string));
            dt.Columns.Add("Sheet", typeof(string));
            dt.Columns.Add("Region", typeof(string));
            dt.Columns.Add("Sample", typeof(string));
            dt.Columns.Add("Image", typeof(Image));
            dt.Columns.Add("Image1", typeof(Image));
            dt.Columns.Add("Image2", typeof(Image));
            dt.Columns.Add("Data", typeof(string));
            dt.Columns.Add("Operator", typeof(string));
            dt.Columns.Add("Time_Update", typeof(string));
            dt.Columns.Add("Remark", typeof(string));
            return dt;
        }
        public static Dictionary<string, List<string>> Get_ProductID(string location, string itemCode, string lotNo)
        {
            if (string.IsNullOrEmpty(location))
            {
                return new Dictionary<string, List<string>>();
            }
            Dictionary<string, List<string>> list = new Dictionary<string, List<string>>();
            ProductIDService service = new ProductIDService(itemCode, lotNo, location, new[] { "GAP" }, new[] { itemCode });
            List<string> listZ = service.getListProductID(service._listFile[itemCode]);
            if (listZ.Count >= 10)
            {
                List<string> doc = new List<string>();
                for (int i = 0; i < 5; i++)
                {
                    doc.Add(listZ[i]);
                }
                List<string> tru = new List<string>();
                for (int i = 0; i < 10; i++)
                {
                    tru.Add(listZ[i]);
                }
                list.Add("DOC", doc);
                list.Add("TRU", tru);
            }
            return list;
        }
    }
}
