using Export_FPCA_OK2ship_Auto_System.Repositories;
using Export_FPCA_OK2ship_Auto_System.Services.TDMK_services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Export_FPCA_OK2ship_Auto_System.Services
{
    public class CrossSectionService
    {
        public CrossSectionService()
        {
        }

        public void export_excel_cross_section(ExcelWorksheet ws, string itemcode, string lotno, DataTable Data_tbl, DataTable dt_spec, int count_sample)
        {

            string regionZ = "";
            int iz = 1;
            foreach (DataRow row in Data_tbl.Rows)
            {
                if (row["Region"].ToString().Equals(regionZ))
                {

                }
                else
                {
                    regionZ = row["Region"].ToString();
                    iz = 1;
                }
                row["Sample"] = iz++;
            }
            List<string> list = new List<string>();
            int counting = int.Parse(dt_spec.Rows[0]["Count_Sample"].ToString());
            for (int i = 0; i < counting; i++)
            {
                list.Add($"Sample {i + 1}");
            }
            list.Add("Flex SN");
            string address = "";
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(ws, list.ToArray(), true);
            foreach (DataRow row in Data_tbl.Rows)
            {
                string dif = "";
                string region = row["Region"].ToString();
                int samplePcs = int.Parse(row["Sample"].ToString()) % counting;
                int step = (int.Parse(row["Sample"].ToString()) - 1) / counting;
                if (samplePcs == 0)
                {
                    samplePcs = counting;
                }
                if (region.Contains("TRAI"))
                {
                    region = region.Replace("TRAI", "T");
                }
                if (region.Contains("PHAI"))
                {
                    region = region.Replace("PHAI", "P");
                }
                switch (region)
                {
                    case "NGANG":
                        if (dic.TryGetValue("Flex SN", out address))
                        {
                            if (dic.TryGetValue($"Sample {samplePcs}", out string addSample))
                            {
                                int vari = 0;
                                if (step <= 0)
                                {
                                    vari = 0;
                                }
                                if (step >= 1)
                                {
                                    vari = 1;
                                }
                                address = ws.Cells[ws.Cells[address.Split('-')[vari]].End.Row, ws.Cells[addSample].End.Column].Address;
                                if (Data_tbl.Columns.Contains("ProductID"))
                                {
                                    ws.Cells[address].Value = row["ProductID"];
                                }
                                address = ExportProcess.AddRow(address, 2 - vari);
                                ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])row[$"Image1"], $"NGANG1{Guid.NewGuid()}");
                                address = ExportProcess.AddRow(address, 1);
                                ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])row[$"Image2"], $"NGANG2{Guid.NewGuid()}");
                                try
                                {
                                    string[] data = row["Data"].ToString().Replace(" ", "").TrimEnd(';').Split(';');
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[0]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[1]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                }
                                catch
                                {
                                    MessageBox.Show($"Có lỗi tại row {row["ID"]}");
                                }
                            }
                        }
                        break;
                    case "TRUNGANG_P":
                    case "TRUNGANG_T":
                        dif = region.Equals("TRUNGANG_P") ? "2 - 5" : "3 - 4";

                        if (dic.TryGetValue("Flex SN", out address))
                        {
                            int z = int.Parse(dif.Split('-')[step]);
                            address = address.Split('-')[z];
                            if (dic.TryGetValue($"Sample {samplePcs}", out string addSample))
                            {
                                address = ws.Cells[ws.Cells[address].End.Row, ws.Cells[addSample].End.Column].Address;
                                if (Data_tbl.Columns.Contains("ProductID"))
                                {
                                    ws.Cells[address].Value = row["ProductID"];
                                }

                                address = ExportProcess.AddRow(address, 1);
                                ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])row["Image1"], $"{Guid.NewGuid()}");
                                try
                                {
                                    string[] data = row["Data"].ToString().Replace(" ", "").TrimEnd(';').Split(';');
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[0]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[1]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                }
                                catch
                                {
                                    MessageBox.Show($"Có lỗi tại row {row["ID"]}");
                                }

                            }
                        }

                        break;
                    case "TRU_T":
                    case "TRU_P":
                    case "DOC_T":
                    case "DOC_P":
                        switch (region)
                        {
                            case "TRU_T":
                                step = 6;
                                break;
                            case "TRU_P":
                                step = 7;
                                break;
                            case "DOC_T":
                                step = 8;
                                break;
                            case "DOC_P":
                                step = 9;
                                break;
                        }
                        if (dic.TryGetValue("Flex SN", out address))
                        {
                            address = address.Split('-')[step];
                            if (dic.TryGetValue($"Sample {samplePcs}", out string addSample))
                            {
                                address = ws.Cells[ws.Cells[address].End.Row, ws.Cells[addSample].End.Column].Address;
                                if (Data_tbl.Columns.Contains("ProductID"))
                                {
                                    ws.Cells[address].Value = row["ProductID"].ToString();
                                }
                                address = ExportProcess.AddRow(address, 1);
                                ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])row["Image1"], $"{Guid.NewGuid()}");
                                address = ExportProcess.AddRow(address, 1);
                                ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])row["Image2"], $"{Guid.NewGuid()}");
                                try
                                {
                                    string[] data = row["Data"].ToString().Replace(" ", "").Split('/');
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[0].TrimEnd(';').Split(';')[0]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[0].TrimEnd(';').Split(';')[1]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[1].TrimEnd(';').Split(';')[1]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[1].TrimEnd(';').Split(';')[4]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[1].TrimEnd(';').Split(';')[2]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[1].TrimEnd(';').Split(';')[5]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[1].TrimEnd(';').Split(';')[0]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                    address = ExportProcess.AddRow(address, 1);
                                    ws.Cells[address].Value = double.Parse(data[1].TrimEnd(';').Split(';')[3]);
                                    ws.Cells[address].Style.Numberformat.Format = "#.##0";
                                }
                                catch
                                {
                                    MessageBox.Show($"Có lỗi tại row {row["ID"]}");
                                }
                            }
                        }
                        break;
                    default:
                        break;
                }

            }
        }

        public string Export(string itemCode, string lotNo)
        {
            try
            {
                using (ExportProcess export = new ExportProcess())
                {
                    using (ExcelPackage package = export.FindFormatProcess("Cross Section", itemCode, lotNo))
                    {
                        using (ExcelWorksheet ws = export.FindSheet(package, "Cross Section"))
                        {

                            int counting;
                            DataTable data_tbl = LoadDT(itemCode, lotNo);
                            DataTable dt_spec = LoadSpec(itemCode, lotNo, out counting);
                            export_excel_cross_section(ws, itemCode, lotNo, data_tbl, dt_spec, counting);
                            export.SaveExcelWorksheet(package, "Cross Section", $"{itemCode}_{lotNo}", "NPI", false);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return "OK";
        }
        private DataTable LoadDT(string itemCode, string lotNo)
        {
            DBContext _dbContext = new DBContext();
            DataTable dataTable = _dbContext.LoadDataTable("CROSS_SECTION_NAS", new[] { "ItemCode", "LotNo", "Sheet" }, new[] { itemCode.PadRight(10), lotNo.PadRight(10), "NPI" });
            if (dataTable.Rows.Count <= 0)
            {
                throw new Exception("No Data");
            }
            string location = dataTable.Rows[0]["LocationImg"].ToString();
            string json = dataTable.Rows[0]["Data"].ToString();
            dataTable = TDMK_ConverterService.JsonToDataTable(json);
            new NasRepository().MergeDataTable(dataTable, "CROSS_SECTION", itemCode, lotNo, location);
            try
            {
                ProductIDService.FillProductID(dataTable, itemCode, lotNo, "CROSS_SECTION");
            }
            catch
            {

            }
            return dataTable;
        }
        private DataTable LoadSpec(string itemCode, string lotNo, out int counting)
        {
            DBContext _dbContext = new DBContext();
            DataTable dataTable = _dbContext.LoadDataTable("SPEC_COMMENT_3", new[] { "ItemCode", "Sheet", "Remark" }, new[] { itemCode.PadRight(10), "CROSS_SECTION", "NPI" });
            counting = int.Parse(dataTable.Rows[0]["Count_Sample"].ToString());
            return dataTable;
        }
    }
}
