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
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Export_FPCA_OK2ship_Auto_System.Services
{
    public class CrossSectionService
    {
        public CrossSectionService()
        {
        }

        public void export_excel_cross_section(ExcelWorksheet ws, string itemcode, string lotno, DataTable Data_tbl,
            DataTable dt_spec, int count_sample)
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

                                address = ws.Cells[ws.Cells[address.Split('-')[vari]].End.Row,
                                    ws.Cells[addSample].End.Column].Address;
                                if (Data_tbl.Columns.Contains("ProductID"))
                                {
                                    ws.Cells[address].Value = row["ProductID"];
                                }

                                address = ExportProcess.AddRow(address, 2 - vari);
                                ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])row[$"Image1"],
                                    $"NGANG1{Guid.NewGuid()}");
                                address = ExportProcess.AddRow(address, 1);
                                ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])row[$"Image2"],
                                    $"NGANG2{Guid.NewGuid()}");
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
                                ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])row["Image1"],
                                    $"{Guid.NewGuid()}");
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
                                ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])row["Image1"],
                                    $"{Guid.NewGuid()}");
                                address = ExportProcess.AddRow(address, 1);
                                ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])row["Image2"],
                                    $"{Guid.NewGuid()}");
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

        public string ExportOld(string itemCode, string lotNo)
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

        public static DataTable load(string itemcode, string lotno, string sheet)
        {
            DBContext _context = new DBContext();
            DataTable dt = _context.LoadDataTable(_DBNAME, new[] { "ItemCode", "LotNo", "Sheet" },
                new[] { itemcode, lotno, sheet });
            return dt;
        }

        public static string _DBNAME = "CROSS_SECTION_NAS";

        public static DataTable HandleData(DataTable dt)
        {
            string locationImage = dt.Rows[0]["LocationImg"].ToString();
            string json = dt.Rows[0]["Data"].ToString();
            string itemCode = dt.Rows[0]["ItemCode"].ToString().Trim();
            string lotNo = dt.Rows[0]["LotNo"].ToString().Trim();
            string sheet = dt.Rows[0]["Sheet"].ToString().Trim();
            DataTable res = TDMK_ConverterService.JsonToDataTable(json);
            NasRepository nas = new NasRepository();
            nas.MergeDataTable(res, _DBNAME, itemCode, $"{lotNo}-{sheet}", locationImage);
            return res;
        }

        private static void ExportWSClip(ExcelWorksheet worksheet, DataTable dtData)
        {
            int maxRow = worksheet.Dimension != null ? worksheet.Dimension.End.Row : 0;

            IDictionary<string, string> dic =
                ExportProcess.FindAddressByText(worksheet,
                    new[] { "Clip Horizontal", "Clip Vertical" });
            foreach (string key in dic.Keys)
            {
                string category = "";
                if (key.ToLower().Contains("horizontal"))
                {
                    category = "DOC";
                }

                if (key.ToLower().Contains("vertical"))
                {
                    category = "NGANG";
                }

                List<DataRow> filteredRows = dtData.AsEnumerable()
                    .Where(row => row.Field<string>("Region").Equals(category))
                    .ToList();
                ExportInOne(worksheet, category, ExportProcess.AddColumn(dic[key], 1), filteredRows, maxRow);
            }
        }

        public static string ExportAll(ExcelPackage package, string itemCode, string lotNo)
        {
            // Load data
            // Get From database
            DBContext _context = new DBContext();
            DataTable dt = _context.LoadDataTable(_DBNAME, new[] { "ItemCode", "LotNo" },
                new[] { itemCode, lotNo });

            // kiểm tra dữ liệu xem có tồn tại hay không
            if (dt.Rows.Count <= 0)
            {
                return "Dữ liệu không tồn tại!";
            }


            string category = dt.Rows[0]["Sheet"].ToString().Trim();
            DataTable dt_data = HandleData(dt);
            try
            {
                if (category.ToUpper().Trim() == "CLIP")
                {
                    return ExportClip(package, dt_data, itemCode, lotNo);
                }

                else if (category.ToUpper().Trim() == "SHIELD B2B")
                {
                    return Export(package, dt_data, itemCode, lotNo);
                }

                else
                {
                    
                }
            }
            catch (Exception ex)
            {
                return $"Error {ex.Message}";
            }

            return "Có vấn đề";
        }

        public static string ExportClip(ExcelPackage package, DataTable dt_data, string itemcode, string lotno)
        {
            List<string> distinctRegions = dt_data.AsEnumerable()
                .Select(row => row.Field<string>("Region"))
                .Distinct()
                .ToList();
            ExportProcess process = new ExportProcess();
            using (ExcelWorksheet worksheet = process.FindSheet(package, "Cross Section Clip"))
            {
                ExportWSClip(worksheet, dt_data);

                return process.SaveExcelPackage(package, $"Cross Section", "");
            }
        }

        public static string Export(ExcelPackage package, DataTable dt_data, string itemcode, string lotno)
        {
            List<string> distinctRegions = dt_data.AsEnumerable()
                .Select(row => row.Field<string>("Region"))
                .Distinct()
                .ToList();

            ExportProcess process = new ExportProcess();
            using (ExcelWorksheet worksheet1 = process.FindSheet(package, "Cross Section (1.1)"))
            using (ExcelWorksheet worksheet2 = process.FindSheet(package, "Cross Section (1.2)"))
            {
                ExportWS(worksheet1, dt_data, distinctRegions);
                ExportWS(worksheet2, dt_data, distinctRegions);
                return process.SaveExcelPackage(package, "Cross Section (1.1):Cross Section (1.2)",
                    $"Cross Section");
            }
        }

        public static void ExportWS(ExcelWorksheet ws, DataTable dt_data, List<string> distinctRegions)
        {
            IDictionary<string, string> dicAddress =
                ExportProcess.FindAddressByText(ws, distinctRegions.ToArray());
            int maxRow = ws.Dimension != null ? ws.Dimension.End.Row : 0;
            foreach (string key in dicAddress.Keys)
            {
                string z = key;

                List<DataRow> filteredRows = dt_data.AsEnumerable()
                    .Where(row => row.Field<string>("Region").Equals(z))
                    .ToList();
                ExportInOne(ws, key, dicAddress[key], filteredRows, maxRow);
            }
        }

        public static void ExportInOne(ExcelWorksheet ws, string category, string address, List<DataRow> filteredRows,
            int maxrow)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            string name = category;
            string addressI = address;
            int i = 0, percent = 0, mma = 0, z = 0;
            int currentRow = int.Parse(Regex.Match(addressI, @"\d+").Value);
            // mainIMG, %percenti, numz 
            while (name.ToUpper() != "FLEX SN" && currentRow <= maxrow)
            {
                currentRow++;
                if (name.Contains(category))
                {
                    dic.Add("mainIMG", addressI);
                }

                else if (name.Contains("%"))
                {
                    dic.Add($"%{percent++}", addressI);
                }

                else if (name.ToUpper().Contains("MIN") || name.ToUpper().Contains("MAX") ||
                         name.ToUpper().Contains("AVERAGE"))
                {
                    dic.Add($"mma{z}_{mma++}", addressI);
                }
                else if (name.Contains("(") && name.Contains(")"))
                {
                    dic.Add($"text{z}_{i++}", addressI);
                }
                else
                {
                    i = 0;
                    mma = 0;
                    dic.Add($"image{z++}", addressI);
                }

                addressI = ExportProcess.AddRow(addressI, 1);
                name = ws.Cells[addressI].Text;
            }

            foreach (DataRow row in filteredRows)
            {
                string nameImage = row["NameImage"].ToString().Split('.')[0];
                string[] split = nameImage.Split('-');
                if (int.TryParse(row["Sample"].ToString(), out int number))
                {
                    if (split.Length > 1)
                    {
                        if (int.TryParse(split[1], out int num))
                        {
                            if (dic.TryGetValue($"image{num - 1}", out string add))
                            {
                                try
                                {
                                    ExportProcess.InsertImageToCell(ws,
                                        ws.Cells[ExportProcess.AddColumn(add, number + 3)], (byte[])row["Image"],
                                        $"{row["NameImage"]}_{number}{Guid.NewGuid()}");
                                }
                                catch (Exception ex)
                                {
                                    Debugger.Break();
                                }
                            }

                            int ix = 0;
                            while (dic.TryGetValue($"%{ix}", out string addPercent))
                            {
                                string targetCell = ExportProcess.AddColumn(addPercent, number + 3);

                                ws.Cells[targetCell].Value = 1;

                                ws.Cells[targetCell].Style.Numberformat.Format = "0%";
                                ix++;
                            }

                            ix = 0;


                            string[] splitData = row["Data"].ToString().Split(';');
                            while (dic.TryGetValue($"text{num}_{ix}", out string addPercent))
                            {
                                if (ix >= splitData.Length)
                                {
                                    break;
                                }

// Đặt biến cho địa chỉ ô để code dễ đọc hơn
                                string targetCell = ExportProcess.AddColumn(addPercent, number + 3);

// Thử chuyển đổi chuỗi thành số thập phân (double)
                                if (double.TryParse(splitData[ix], out double parsedNumber))
                                {
                                    // Làm tròn 2 chữ số và gán vào ô
                                    ws.Cells[targetCell].Value = Math.Round(parsedNumber, 2);

                                    // (Tùy chọn) Định dạng ô Excel để luôn hiển thị đủ 2 chữ số (VD: hiển thị 15.50 thay vì 15.5)
                                    ws.Cells[targetCell].Style.Numberformat.Format = "0.00";
                                }
                                else
                                {
                                    // Nếu dữ liệu không phải là số (ví dụ: bị rỗng, hoặc chữ "N/A"), cứ ghi chuỗi gốc vào
                                    ws.Cells[targetCell].Value = "N/A";
                                }

                                ix++;
                            }

                            while (dic.TryGetValue($"mma{num}_{ix}", out string addPercent))
                            {
                                int ireal;
                                switch (ix)
                                {
                                    // ỉneal
                                    case 0:
                                        ireal = 1;
                                        break;
                                    case 1:
                                        ireal = 4;
                                        break;

                                    case 2:
                                        ireal = 2;
                                        break;

                                    case 3:
                                        ireal = 5;
                                        break;

                                    case 4:
                                        ireal = 0;
                                        break;

                                    case 5:
                                        ireal = 3;
                                        break;
                                    default:
                                        ireal = ix;
                                        break;
                                }

                                if (ireal < splitData.Length)
                                {
                                    string targetCell = ExportProcess.AddColumn(addPercent, number + 3);

                                    // Thử chuyển đổi chuỗi thành số thập phân, dùng InvariantCulture để tránh lỗi dấu phẩy/chấm
                                    if (double.TryParse(splitData[ireal], System.Globalization.NumberStyles.Any,
                                            System.Globalization.CultureInfo.InvariantCulture,
                                            out double parsedNumberReal))
                                    {
                                        // Gán giá trị đã làm tròn 2 chữ số
                                        ws.Cells[targetCell].Value = Math.Round(parsedNumberReal, 2);

                                        // Định dạng ô trên Excel để hiển thị chuẩn 2 số thập phân (VD: 10.00)
                                        ws.Cells[targetCell].Style.Numberformat.Format = "0.00";
                                    }
                                    else
                                    {
                                        // Gán lại chuỗi gốc nếu không thể chuyển thành số
                                        ws.Cells[targetCell].Value = "N/A";
                                    }
                                }

                                ix++;
                            }
                        }
                    }
                    else
                    {
                        if (dic.TryGetValue("mainIMG", out string add))
                        {
                            try
                            {
                                int ix = 0;
                                ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(add, number + 3), -1)]
                                    .Value = row["ProductID"].ToString();
                                ExportProcess.InsertImageToCell(ws, ws.Cells[ExportProcess.AddColumn(add, number + 3)],
                                    (byte[])row["Image"], $"{row["NameImage"]}_{number}{Guid.NewGuid()}");
                                string[] splitData = row["Data"].ToString().Split(';');
                                while (dic.TryGetValue($"text0_{ix}", out string addPercent))
                                {
                                    if (ix >= splitData.Length)
                                    {
                                        break;
                                    }

// Đặt biến cho địa chỉ ô để code dễ đọc hơn
                                    string targetCell = ExportProcess.AddColumn(addPercent, number + 3);

// Thử chuyển đổi chuỗi thành số thập phân (double)
                                    if (double.TryParse(splitData[ix], out double parsedNumber))
                                    {
                                        // Làm tròn 2 chữ số và gán vào ô
                                        ws.Cells[targetCell].Value = Math.Round(parsedNumber, 2);

                                        // (Tùy chọn) Định dạng ô Excel để luôn hiển thị đủ 2 chữ số (VD: hiển thị 15.50 thay vì 15.5)
                                        ws.Cells[targetCell].Style.Numberformat.Format = "0.00";
                                    }
                                    else
                                    {
                                        // Nếu dữ liệu không phải là số (ví dụ: bị rỗng, hoặc chữ "N/A"), cứ ghi chuỗi gốc vào
                                        ws.Cells[targetCell].Value = "N/A";
                                    }

                                    ix++;
                                }
                            }
                            catch (Exception ex)
                            {
                                Debugger.Break();
                            }
                        }
                    }
                }
            }

            ws.Cells[address].Value = "";
        }

        private DataTable LoadDT(string itemCode, string lotNo)
        {
            DBContext _dbContext = new DBContext();
            DataTable dataTable = _dbContext.LoadDataTable("CROSS_SECTION_NAS", new[] { "ItemCode", "LotNo", "Sheet" },
                new[] { itemCode.PadRight(10), lotNo.PadRight(10), "NPI" });
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
            DataTable dataTable = _dbContext.LoadDataTable("SPEC_COMMENT_3", new[] { "ItemCode", "Sheet", "Remark" },
                new[] { itemCode.PadRight(10), "CROSS_SECTION", "NPI" });
            counting = int.Parse(dataTable.Rows[0]["Count_Sample"].ToString());
            return dataTable;
        }
    }
}