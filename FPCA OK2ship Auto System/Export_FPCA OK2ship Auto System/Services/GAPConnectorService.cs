using Export_FPCA_OK2ship_Auto_System.Repositories;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Export_FPCA_OK2ship_Auto_System.Services
{
    public class GAPConnectorService
    {
        public void Export(ExcelWorksheet workSheet, DataTable dataTable, DataTable spec)
        {

            int SpecNum = int.Parse(spec.Rows[0]["Count_Sample"].ToString());
            string[] healder = new[] { "IO PIN", "GROUNDING PIN_LEFT", "GROUNDING PIN_RIGHT", "Sample 1" };
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
                                                  .Where(row => row.Field<string>("region") == "GAP DOC")
                                                .CopyToDataTable();
                        break;
                    case "GROUNDING PIN_RIGHT":
                        originalTable = dataTable.AsEnumerable()
                                                  .Where(row => row.Field<string>("region") == "GAP DOC")
                                                .CopyToDataTable();
                        break;

                }
                bool prime = false;
                string address = _heal[dic[item]];
                for (int i = 0; i < SpecNum; i++)
                {
                    string addressP = ExportProcess.AddColumn(address, i);
                    addressP = ExportProcess.AddRow(addressP, 1);
                    // Kiểm tra và ghi product ID
                    if (prime || workSheet.Cells[ExportProcess.AddColumn(addressP, -1)].Text.Contains("Flex"))
                    {
                        prime = true;
                        if (originalTable.Columns.Contains("ProductID"))
                        {
                            workSheet.Cells[addressP].Value = originalTable.Rows[i]["ProductID"];
                        }
                    }
                    else
                    {
                        addressP = ExportProcess.AddRow(addressP, -1);
                    }

                    addressP = ExportProcess.AddRow(addressP, 1);
                    ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressP], (byte[])originalTable.Rows[i]["Image"], $"{Guid.NewGuid()}");
                    addressP = ExportProcess.AddRow(addressP, 1);
                    ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressP], (byte[])originalTable.Rows[i]["Image1"], $"{Guid.NewGuid()}");
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
                    ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressP], (byte[])originalTable.Rows[i]["Image2"], $"{Guid.NewGuid()}");
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

        }
      
    }
}
