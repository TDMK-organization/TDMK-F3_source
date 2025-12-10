using Export_FPCA_OK2ship_Auto_System.Repositories;
using Export_FPCA_OK2ship_Auto_System.Services.TDMK_services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Export_FPCA_OK2ship_Auto_System.Services
{
    public class PeelPullShearService
    {
        public void export_excel_peel_pull_shear(ExcelWorksheet ws, string itemcode, string lotno, string sheet, DataTable Data_tbl, DataTable dt_spec, int count_sample, bool prime = false)
        {


            string testName = "Force!!!!!!!!!!!!!!!", forceName = "Froezzzz", caching = "crack";
            if (ws.Name.Contains("Peel"))
            {
                testName = "Peeling Force";
                forceName = "Judgement peeling force";
            }
            if (ws.Name.Contains("Pull"))
            {
                testName = "Pulling Force";
                forceName = "Judgement pulling force";
            }
            if (ws.Name.Contains("Shear"))
            {
                testName = "Shear Force (Kgf)";
                forceName = "Judgement Shear force";
                caching = "shear";
            }


            string[] nameAddress = new[] { "Flex SN", testName, forceName, "Sample", "STDEV", "CPK", "Min Force (N)", "Max Force (N)", "Average Force (N)", "Picture", "Graph", "Final judgement", "Judgement failure mode", $"Solder joint {caching}", "Pad lift", "Solder joint lift", "Intermetallic break", "Component damage", "Component detached", "Flex torn" };
            //Find Dictionary have name column {sample 1} and address
            IDictionary<string, string> addressDic = ExportProcess.FindAddressByText(ws, nameAddress);

            string addressForce = "";
            if (addressDic.TryGetValue("Sample", out string address))
            {
                if (address.Split('-').Count() < Data_tbl.Rows.Count)
                {
                    int z = Data_tbl.Rows.Count - address.Split('-').Count();
                    string a = address.Split('-')[address.Split('-').Count() - 1];
                    for (int c = 0; c < z; c++)
                    {
                        string aStart = ExportProcess.AddColumn(ExportProcess.AddRow(a, -1), c);
                        string aEnd = ExportProcess.AddColumn(ExportProcess.AddRow(a, 19), c);
                        ExportProcess.CopyColumn(ws, ws.Cells[$"{aStart}:{aEnd}"], ExportProcess.AddRow(ExportProcess.AddColumn(a, c + 1), -1));
                        ws.Cells[ExportProcess.AddColumn(a, c + 1)].Value = $"Sample {count_sample + c + 1}";
                        addressDic["Sample"] = $"{addressDic["Sample"]}-{ExportProcess.AddColumn(a, c + 1)}";
                    }
                    address = addressDic["Sample"];
                }
                if (count_sample != address.Split('-').Count())
                {
                    //if (MessageBox.Show("Số pcs của format không phù hợp bạn có muốn tiếp tục", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.No)
                    //{
                    //    return;
                    //}

                }

            }

            try
            {
                int rowIndex = -1;
                IList<double> numbers = new List<double>();
                foreach (string addressCol in address.Split('-'))
                {
                    rowIndex++;
                    int column = ws.Cells[addressCol].End.Column;
                    ///
                    DataRow item = Data_tbl.Rows[rowIndex];
                    if (addressDic.TryGetValue("Flex SN", out string adz))
                    {
                        adz = ws.Cells[ws.Cells[adz].Start.Row, ws.Cells[addressCol].Start.Column].Address;
                        if (Data_tbl.Columns.Contains("ProductID"))
                        {
                            ws.Cells[adz].Value = item["ProductID"];
                        }
                    }
                    if (addressDic.TryGetValue("Picture", out string addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        if (item["Image"] != DBNull.Value && item["Image"] is byte[])
                        {
                            ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])item["Image"], $"{rowIndex} - picture");
                        }
                        if (item["Image"] != DBNull.Value && item["Image"] is Image)
                        {
                            ExportProcess.InsertImageToCell(ws, ws.Cells[address], TDMK_ImageConverter.ImageToByteArray((Image)item["Image"], ImageFormat.Jpeg), $"{rowIndex} - picture");
                        }
                    }
                    if (addressDic.TryGetValue("Graph", out addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        if (item["Graph"] != DBNull.Value && item["Graph"] is byte[])
                        {
                            ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])item["Graph"], $"{rowIndex} - Graph");
                        }
                        if (item["Graph"] != DBNull.Value && item["Graph"] is Image)
                        {
                            ExportProcess.InsertImageToCell(ws, ws.Cells[address], TDMK_ImageConverter.ImageToByteArray((Image)item["Graph"], ImageFormat.Jpeg), $"{rowIndex} - Graph");
                        }
                    }
                    if (addressDic.TryGetValue(testName, out addressRow))
                    {
                        addressRow = addressRow.Split('-')[0];
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        int z = testName.Contains("Shear") ? 1 : 0;
                        addressForce += $" {ExportProcess.AddRow(address, z)} + ";
                        ws.Cells[address].Value = double.Parse(item["Data"].ToString().Trim());
                        if (testName.Contains("Shear"))
                        {
                            ws.Cells[ExportProcess.AddRow(address, 1)].FormulaR1C1 = $" =R[-1]C*9.8";
                        }
                    }


                    string primeMode = "";
                    double maxMode = double.MinValue;
                    int pin = 0;
                    try
                    {
                        string pinZ = item["Mode 1: Solder joint crack"].ToString().Split('%')[1].Split('/')[1].Replace(")", "");
                        pin = int.Parse(pinZ);
                    }
                    catch
                    {

                    }
                    if (addressDic.TryGetValue($"Solder joint {caching}", out addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        string[] content = item["Mode 1: Solder joint crack"].ToString().Split('%');
                        if (content.Count() > 1 && content[1].Contains("("))
                        {
                            ws.Cells[address].Formula = $"={content[1]}";

                            if (double.TryParse(content[1].Split('/')[0].Replace("(", ""), out double number))
                            {
                                if (number > maxMode)
                                {
                                    maxMode = number;
                                    primeMode = "Mode #1";
                                }
                            }
                        }
                        else
                        {
                            ws.Cells[address].Formula = $"=0/{pin}";
                        }
                    }
                    if (addressDic.TryGetValue("Pad lift", out addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        string[] content = item["Mode 2: Pad lift"].ToString().Split('%');
                        if (content.Count() > 1 && content[1].Contains("("))
                        {
                            ws.Cells[address].Formula = $"={content[1]}";

                            if (double.TryParse(content[1].Split('/')[0].Replace("(", ""), out double number))
                            {
                                if (number > maxMode)
                                {
                                    maxMode = number;
                                    primeMode = "Mode #2";
                                }
                            }
                        }
                        else
                        {
                            ws.Cells[address].Formula = $"=0/{pin}";
                        }
                    }
                    if (addressDic.TryGetValue("Solder joint lift", out addressRow))
                    {

                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        string[] content = item["Mode 3: Solder joint lift"].ToString().Split('%');
                        if (content.Count() > 1 && content[1].Contains("("))
                        {
                            ws.Cells[address].Formula = $"={content[1]}";

                            if (double.TryParse(content[1].Split('/')[0].Replace("(", ""), out double number))
                            {
                                if (number > maxMode)
                                {
                                    maxMode = number;
                                    primeMode = "Mode #3";
                                }
                            }
                        }
                        else
                        {
                            ws.Cells[address].Formula = $"=0/{pin}";
                        }
                    }
                    if (addressDic.TryGetValue("Intermetallic break", out addressRow))
                    {

                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        string[] content = item["Mode 4: Intermetallic break"].ToString().Split('%');
                        if (content.Count() > 1 && content[1].Contains("("))
                        {
                            ws.Cells[address].Formula = $"={content[1]}";

                            if (double.TryParse(content[1].Split('/')[0].Replace("(", ""), out double number))
                            {
                                if (number > maxMode)
                                {
                                    maxMode = number;
                                    primeMode = "Mode #4";
                                }
                            }
                        }
                        else
                        {
                            ws.Cells[address].Formula = $"=0/{pin}";
                        }
                    }
                    if (addressDic.TryGetValue("Component damage", out addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        string[] content = item["Mode 5: Component damage"].ToString().Split('%');
                        if (content.Count() > 1 && content[1].Contains("("))
                        {
                            ws.Cells[address].Formula = $"={content[1]}";

                            if (double.TryParse(content[1].Split('/')[0].Replace("(", ""), out double number))
                            {
                                if (number > maxMode)
                                {
                                    maxMode = number;
                                    primeMode = "Mode #5";
                                }
                            }
                        }
                        else
                        {
                            ws.Cells[address].Formula = $"=0/{pin}";
                        }
                    }
                    if (addressDic.TryGetValue("Component detached", out addressRow))
                    {

                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        string[] content = item["Mode 6: Component detached"].ToString().Split('%');
                        if (content.Count() > 1 && content[1].Contains("("))
                        {
                            ws.Cells[address].Formula = $"={content[1]}";

                            if (double.TryParse(content[1].Split('/')[0].Replace("(", ""), out double number))
                            {
                                if (number > maxMode)
                                {
                                    maxMode = number;
                                    primeMode = "Mode #6";
                                }
                            }
                        }
                        else
                        {
                            ws.Cells[address].Formula = $"=0/{pin}";
                        }
                    }
                    if (addressDic.TryGetValue("Flex torn", out addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        string[] content = item["Mode 7: Flex torn"].ToString().Split('%');
                        if (content.Count() > 1 && content[1].Contains("("))
                        {
                            ws.Cells[address].Formula = $"={content[1]}";

                            if (double.TryParse(content[1].Split('/')[0].Replace("(", ""), out double number))
                            {
                                if (number > maxMode)
                                {
                                    maxMode = number;
                                    primeMode = "Mode #7";
                                }
                            }
                        }
                        else
                        {
                            ws.Cells[address].Formula = $"=0/{pin}";
                        }
                    }
                    //Debugger.Break();
                    double valueForce = double.MinValue;

                    if (addressDic.TryGetValue(forceName, out addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;

                        if (double.TryParse(item["Data"].ToString(), out valueForce))
                        {
                            if (dt_spec.Rows.Count > 0)
                            {
                                DataRow row = dt_spec.Rows[0];
                                string specLow = row["Location"].ToString().Replace("N", "").Replace(">", "").Trim();
                                if (double.TryParse(specLow, out double num))
                                {
                                    if (addressDic.TryGetValue(testName, out string addressRowZ))
                                    {
                                        addressRowZ = addressRowZ.Split('-').FirstOrDefault();
                                        int r = ws.Cells[address].Start.Row - ws.Cells[addressRowZ].Start.Row;
                                        ws.Cells[address].FormulaR1C1 = $"=IF(R[{r * -1}]C>{num}, \"Pass\",\"Fail\" )";
                                    }

                                }
                            }
                            else
                            {
                                ws.Cells[address].Value = valueForce > 5 ? "Pass" : "NG";
                            }
                        }
                        numbers.Add(valueForce);
                    }
                    if (addressDic.TryGetValue("Judgement failure mode", out addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        //ws.Cells[address].Value = primeMode;
                        if (addressDic.TryGetValue("Flex torn", out string addMode7))
                        {
                            string addressMode7 = ws.Cells[ws.Cells[addMode7].End.Row, column].Address;
                            if (addressDic.TryGetValue($"Solder joint {caching}", out string addMode1))
                            {
                                string addressMode1 = ws.Cells[ws.Cells[addMode1].End.Row, column].Address;

                                addMode1 = ExportProcess.AddColumn(addMode1, -1);
                                addMode7 = ExportProcess.AddColumn(addMode7, -1);
                                int r1 = ws.Cells[addMode1].End.Row - ws.Cells[address].End.Row;
                                int c1 = ws.Cells[addMode1].End.Column - ws.Cells[address].End.Column;
                                int r7 = ws.Cells[addMode7].End.Row - ws.Cells[address].End.Row;
                                int c7 = ws.Cells[addMode7].End.Column - ws.Cells[address].End.Column;


                                int row1 = ws.Cells[addressMode1].End.Row - ws.Cells[address].End.Row;
                                int col1 = ws.Cells[addressMode1].End.Column - ws.Cells[address].End.Column;
                                int row7 = ws.Cells[addressMode7].End.Row - ws.Cells[address].End.Row;
                                int col7 = ws.Cells[addressMode7].End.Column - ws.Cells[address].End.Column;

                                ws.Cells[address].FormulaR1C1 = $"=INDEX(R[{r1}]C[{c1}]:R[{r7}]C[{c7}],MATCH(MAX(R[{row1}]C[{col1}]:R[{row7}]C[{col7}]),R[{row1}]C[{col1}]:R[{row7}]C[{col7}],0))";
                            }
                        }
                    }
                    if (addressDic.TryGetValue("Final judgement", out addressRow))
                    {
                        address = ws.Cells[ws.Cells[addressRow].End.Row, column].Address;
                        if (addressDic.TryGetValue("Intermetallic break", out string addressRowZ))
                        {
                            int r = ws.Cells[addressRowZ].End.Row - ws.Cells[addressRow].End.Row;
                            ws.Cells[address].FormulaR1C1 = $"=IF(R[{r}]C > 0, \"SEM/EDX analysis 1 piece\", \"Pass\")";
                        }

                        //if (valueForce > 5 && (primeMode.Contains("5") || primeMode.Contains("2")))
                        //{
                        //    ws.Cells[address].Value = "Pass";
                        //}
                        //else
                        //{
                        //    ws.Cells[address].Value = "Fail";

                        //}

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            if (addressDic.TryGetValue("Max Force (N)", out string addressRowz))
            {
                address = ws.Cells[ws.Cells[addressRowz].End.Row, ws.Cells[addressRowz].End.Column + 2].Address;
                ws.Cells[address].FormulaR1C1 = $"=MAX({ExportProcess.ConvertAddressRangeBase(ws, addressForce, address)})";
            }
            if (addressDic.TryGetValue("Min Force (N)", out addressRowz))
            {
                address = ws.Cells[ws.Cells[addressRowz].End.Row, ws.Cells[addressRowz].End.Column + 2].Address;
                ws.Cells[address].FormulaR1C1 = $"=MIN({ExportProcess.ConvertAddressRangeBase(ws, addressForce, address)})";
            }
            if (addressDic.TryGetValue("Average Force (N)", out string addressAverage))
            {
                addressAverage = ws.Cells[ws.Cells[addressAverage].End.Row, ws.Cells[addressAverage].End.Column + 2].Address;
                ws.Cells[addressAverage].FormulaR1C1 = $"=average({ExportProcess.ConvertAddressRangeBase(ws, addressForce, addressAverage)})";
            }

            if (addressDic.TryGetValue("STDEV", out string addressSTDEV))
            {
                addressSTDEV = ws.Cells[ws.Cells[addressSTDEV].End.Row, ws.Cells[addressRowz].End.Column + 2].Address;
                ws.Cells[addressSTDEV].FormulaR1C1 = $"=STDEV({ExportProcess.ConvertAddressRangeBase(ws, addressForce, addressSTDEV)})";
                if (addressDic.TryGetValue("CPK", out addressRowz))
                {
                    address = ws.Cells[ws.Cells[addressRowz].End.Row, ws.Cells[addressRowz].End.Column + 2].Address;
                    ws.Cells[address].FormulaR1C1 = $"=({ExportProcess.ConvertAddressRangeBase(ws, addressAverage, address)}-{TDMK_ConverterService.GetNumberFromString(dt_spec.Rows[0]["Location"].ToString())})/(3*{ExportProcess.ConvertAddressRangeBase(ws, addressSTDEV, address)})";
                }
            }

            if (prime)
            {
                ws.Name = "Peel Test without SUS";
            }
        }

        public string Export(string itemCode, string lotNo, string category)
        {
            string sheet = "";
            switch (category)
            {
                case "Peel Test":
                    sheet = "Peel Test";
                    break;
                case "(Mating) Pull Test":
                    sheet = "(Mating) Pull Test";
                    break;
                case "Shear test":
                    sheet = "Shear test";
                    break;
                default:
                    return "Category không hợp lệ";
            }
            try
            {
                using (ExportProcess export = new ExportProcess())
                {
                    using (ExcelPackage package = export.FindFormatProcess(sheet, itemCode, lotNo))
                    {
                        using (ExcelWorksheet ws = export.FindSheet(package, sheet))
                        {
                            int counting;
                            DataTable data_tbl = LoadDT(itemCode, lotNo, sheet);
                            DataTable dt_spec = LoadSpec(itemCode, lotNo, sheet, out counting);
                            export_excel_peel_pull_shear(ws, itemCode, lotNo, sheet, data_tbl, dt_spec, counting, false);
                            export.SaveExcelWorksheet(package, sheet, $"{itemCode}_{lotNo}", "NPI", false);
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

        private DataTable LoadSpec(string itemCode, string lotNo, string category, out int counting)
        {
            string tableName = "";
            switch (category)
            {
                case "Peel Test":
                    tableName = "PEEL_TEST";
                    break;
                case "(Mating) Pull Test":
                    tableName = "MATING_PULL_TEST";
                    break;
                case "Shear test":
                    tableName = "SHEAR_TEST";
                    break;
                default:
                    throw new Exception("Category không hợp lệ");
            }
            DataTable dataTable = new DBContext().LoadDataTable("SPEC_COMMENT_3", new[] { "ItemCode", "Sheet", "Remark" }, new[] { itemCode.PadRight(10), tableName, "NPI" });
            counting = int.Parse(dataTable.Rows[0]["Count_Sample"].ToString());
            return dataTable;

        }

        private DataTable LoadDT(string itemCode, string lotNo, string category)
        {
            string tableName = "";
            switch (category)
            {
                case "Peel Test":
                    tableName = "PEEL_TEST_NAS";
                    break;
                case "(Mating) Pull Test":
                    tableName = "MATING_PULL_TEST_NAS";
                    break;
                case "Shear test":
                    tableName = "SHEAR_TEST_NAS";
                    break;
                default:
                    throw new Exception("Category không hợp lệ");
            }
            DBContext _dbContext = new DBContext();
            DataTable dataTable = _dbContext.LoadDataTable(tableName, new[] { "ItemCode", "LotNo", "Sheet" }, new[] { itemCode.PadRight(10), lotNo.PadRight(10), "NPI" });
            if (dataTable.Rows.Count <= 0)
            {
                throw new Exception("No Data");
            }
            string location = dataTable.Rows[0]["LocationImg"].ToString();
            string json = dataTable.Rows[0]["Data"].ToString();
            dataTable = TDMK_ConverterService.JsonToDataTable(json);
            new NasRepository().MergeDataTable(dataTable, tableName, itemCode, lotNo, location);
            try
            {
                ProductIDService.FillProductID(dataTable, itemCode, lotNo, tableName.Replace("_NAS", ""));
            }
            catch
            {

            }
            return dataTable;
        }
    }
}
