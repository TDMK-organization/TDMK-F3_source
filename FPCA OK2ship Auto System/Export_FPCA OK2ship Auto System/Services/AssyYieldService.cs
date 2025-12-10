using Export_FPCA_OK2ship_Auto_System.Repositories;
using Export_FPCA_OK2ship_Auto_System.Services.TDMK_services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Export_FPCA_OK2ship_Auto_System.Services
{
    public class AssyYieldService
    {
        private DBContext _dbContext = null;
        private string _nameTable = "ASSY_YIELD";
        public AssyYieldService(SqlConnection sqlcon)
        {
            _dbContext = new DBContext(sqlcon);
        }

        public Dictionary<string, DataTable> Load(string itemCode, string lotNo)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            if (string.IsNullOrEmpty(itemCode) && string.IsNullOrEmpty(lotNo))
            {
                throw new Exception("Không được để trống itemCode lotno");
            }

            DataTable data = _dbContext.LoadDataTable(_nameTable, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            if (data.Rows.Count < 1)
            {
                throw new Exception("Không tìm thấy dữ liệu");
            }
            Guid area = Guid.Parse(data.Rows[0]["Area"].ToString());
            DataTable dataImage = _dbContext.LoadDataTable(_nameTable + "_IMAGE", new[] { "Area" }, new[] { area.ToString() });
            string[] healder = new string[] { "Process", "Top_SMT", "Top_Backend", "OQC" };
            Dictionary<string, DataTable> dIC = new Dictionary<string, DataTable>();
            foreach (string item in healder)
            {
                string json = data.Rows[0][item].ToString();

                DataTable zz = TDMK_ConverterService.JsonToDataTable(json);
                if (zz.Rows.Count <= 0)
                {
                    continue;
                }
                else
                {

                    try
                    {
                        DataTable z = TDMK_ConverterService.ConvertDataTableImage(zz, dataImage);
                        dIC.Add(item, z);
                    }
                    catch { }
                }

            }
            return dIC;

        }
        public string Export(string itemCode, string lotNo)
        {
            _dbContext = new DBContext();
            try
            {

                using (ExportProcess process = new ExportProcess())
                {
                    using (ExcelPackage package = process.FindFormatProcess("Assy Yield", itemCode, lotNo))
                    {
                        using (ExcelWorksheet worksheet = process.FindSheet(package, "Assy Yield"))
                        {
                            string msg = Export(worksheet, itemCode, lotNo);
                            process.SaveExcelWorksheet(package, "Assy Yield", $"{itemCode}_{lotNo}", "NPI", false);
                            return msg;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        public void ExportTableOfContent(ExcelWorksheet worksheet, string itemCode, string lotName)
        {
            Debugger.Break();
            DataTable db = _dbContext.LoadDataTable("TABLE_OF_CONTENT_SETTING", new[] { "ItemCode" }, new[] { itemCode });
            if (db.Rows.Count < 1)
            {
                return;
            }
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(worksheet, new[] { "Build Config", "Program Name", "Lot #", "ODB++ & Revision", "MCO & Revision" });
            worksheet.Cells[ExportProcess.AddColumn(dic["Program Name"].ToString(), 1)].Value = db.Rows[0]["ProgramName"];
            worksheet.Cells[ExportProcess.AddColumn(dic["Lot #"].ToString(), 1)].Value = lotName;
            worksheet.Cells[ExportProcess.AddColumn(dic["ODB++ & Revision"].ToString(), 1)].Value = db.Rows[0]["ODBRevision"];
            worksheet.Cells[ExportProcess.AddColumn(dic["MCO & Revision"].ToString(), 1)].Value = db.Rows[0]["MCORevision"];
            worksheet.Cells[ExportProcess.AddColumn(dic["Build Config"].ToString(), 1)].Value = db.Rows[0]["Build"];
        }
        public string Export(ExcelWorksheet worksheet, string itemCode, string lotNo)
        {

            Dictionary<string, DataTable> dIC = Load(itemCode, lotNo);
            ExportProcess exportProcess = new ExportProcess();


            ExportTableOfContent(worksheet, itemCode, lotNo);
            string[] healder = new string[] { "Production Yield Target:", "Station", "Input", "Passed and shipped to next process", "Rejected", "Evaluation", "IPQC", "ORT", "WIP", "Others" };
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(worksheet, healder, true);
            if (dic.TryGetValue("Production Yield Target:", out string valueZA))
            {
                DataTable dtz = _dbContext.LoadDataTable("TARGET_OF_ASSY_YIELD", new[] { "ItemCode" }, new[] { itemCode });
                string value = dtz.Rows[0]["Value"].ToString();
                if (double.TryParse(value, out double val))
                {
                    worksheet.Cells[ExportProcess.AddColumn(valueZA, 1)].Value = val / 100;
                    worksheet.Cells[ExportProcess.AddColumn(valueZA, 1)].Style.Numberformat.Format = "#0.00%";

                }
            }
            #region Process
            DataTable dataTable = dIC["Process"];
            Dictionary<string, int> Marking = new Dictionary<string, int>();
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                DataRow row = dataTable.Rows[i];
                Marking.Add(row["Station"].ToString(), i);
            }
            string address = dic["Station"].Split('-')[dic["Station"].Split('-').Count() - 1];
            string SaveADD = address;
            while (true)
            {
                address = ExportProcess.AddRow(address, 1);
                string value = worksheet.Cells[address].Text;
                if (string.IsNullOrEmpty(value))
                {
                    break;
                }
                if (Marking.TryGetValue(value, out int rowNum))
                {
                    DataRow row = dataTable.Rows[rowNum];
                    int s1 = 0, s2 = int.MaxValue, s3 = int.MinValue;
                    if (dic.TryGetValue("Passed and shipped to next process", out string addRott))
                    {
                        addRott = worksheet.Cells[worksheet.Cells[address].End.Row, worksheet.Cells[addRott].End.Column].Address;
                    }
                    if (dic.TryGetValue("Input", out string addStation))
                    {
                        addStation = worksheet.Cells[worksheet.Cells[address].End.Row, worksheet.Cells[addStation].End.Column].Address;
                        worksheet.Cells[addStation].Value = int.Parse(row["Input"].ToString());
                        s1 = worksheet.Cells[addStation].End.Column - worksheet.Cells[addRott].End.Column;
                    }
                    for (int j = dataTable.Columns.Count - 1; j > 2; j--)
                    {
                        DataColumn column = dataTable.Columns[j];
                        string columnName = column.ColumnName;
                        if (dic.TryGetValue(columnName, out string addressCol))
                        {
                            addressCol = addressCol.Split('-')[addressCol.Split('-').Count() - 1];
                            addressCol = worksheet.Cells[worksheet.Cells[address].End.Row, worksheet.Cells[addressCol].End.Column].Address;
                            worksheet.Cells[addressCol].Value = int.Parse(row[columnName].ToString());
                            int solveInt = worksheet.Cells[addressCol].End.Column - worksheet.Cells[addRott].End.Column;
                            s3 = Math.Max(s3, solveInt);
                            if (solveInt > 0)
                            {
                                s2 = Math.Min(s2, solveInt);
                            }

                        }
                    }
                    worksheet.Cells[addRott].FormulaR1C1 = $"=RC[{s1}]-SUM(RC[{s2}]:RC[{s3}])";
                    //Debugger.Break();
                }

            }
            #endregion
            #region TOP 
            List<string> heaelso = new List<string>();
            foreach (string item in dIC.Keys.ToArray())
            {
                heaelso.Add(item.Replace('_', ' '));
            }
            dic = ExportProcess.FindAddressByText(worksheet, heaelso.ToArray());
            if (dic.TryGetValue("OQC", out string jsonzzz))
            {
                dic["OQC"] = dic["OQC"].Split('-')[dic["OQC"].Split('-').Count() - 1];
            }
            dic.Remove("Process");

            foreach (string item in heaelso)
            {
                if (!item.Equals("Process"))
                {
                    string addressTop = ExportProcess.AddRow(dic[item], 2);
                    dataTable = dIC[item.Replace(' ', '_')];
                    if (dataTable.Rows.Count < 2)
                    {

                    }
                    else
                    {
                        int coluAdd = dataTable.Rows.Count - 1;
                        worksheet.InsertRow(worksheet.Cells[addressTop].End.Row + 1, coluAdd);
                        for (int i = 1; i <= coluAdd; i++)
                        {
                            string addressTopz = worksheet.Cells[worksheet.Cells[addressTop].End.Row + i, 1, worksheet.Cells[addressTop].End.Row + i, 20].Address;
                            addressTop = worksheet.Cells[worksheet.Cells[addressTop].End.Row, 1, worksheet.Cells[addressTop].End.Row, 20].Address;
                            worksheet.Cells[addressTop].Copy(worksheet.Cells[addressTopz]);
                            worksheet.Cells[addressTop].CopyStyles(worksheet.Cells[addressTopz]);
                            worksheet.Row(worksheet.Cells[addressTopz].End.Row).Height = worksheet.Row(worksheet.Cells[addressTop].End.Row).Height;
                        }

                        dic = ExportProcess.FindAddressByText(worksheet, heaelso.ToArray());
                        if (dic.TryGetValue("OQC", out string jsonzszz))
                        {
                            dic["OQC"] = dic["OQC"].Split('-')[dic["OQC"].Split('-').Count() - 1];
                        }
                        dic.Remove("Process");
                    }

                    addressTop = ExportProcess.AddColumn(ExportProcess.AddRow(dic[item], 2), -1);

                    int iz = 0;
                    foreach (DataRow row in dataTable.Rows)
                    {
                        string add = ExportProcess.AddRow(addressTop, iz);
                        foreach (DataColumn col in dataTable.Columns)
                        {
                            if (col.DataType == typeof(byte[]) || col.DataType == typeof(Image))
                            {

                                ExcelRangeBase newz = worksheet.Cells[worksheet.Cells[add].Address];
                                try
                                {

                                    if (row[col] is byte[])
                                    {
                                        ExportProcess.InsertImageToCell(worksheet, newz, (byte[])row[col], $"{Guid.NewGuid()}");
                                    }
                                    else
                                    {
                                        ExportProcess.InsertImageToCell(worksheet, newz, TDMK_ImageConverter.ImageToByteArray((Image)row[col], ImageFormat.Png), $"{Guid.NewGuid()}");
                                    }
                                }
                                catch
                                {

                                }
                                newz.Value = row["Defect Name"];
                                add = ExportProcess.AddColumn(add, 1);

                            }
                            else
                            {
                                if (col.ColumnName.Equals("Defect Name"))
                                {
                                    add = ExportProcess.AddColumn(add, -1);

                                }
                                if (col.ColumnName.Equals("Defect Rate"))
                                {

                                    int r = worksheet.Cells[SaveADD].Start.Row - worksheet.Cells[add].Start.Row + 1;
                                    int c = worksheet.Cells[SaveADD].Start.Column - worksheet.Cells[add].Start.Column + 1;
                                    worksheet.Cells[add].FormulaR1C1 = $"=RC[-1]/R[{r}]C[{c}]";
                                }
                                else
                                {
                                    worksheet.Cells[add].Value = row[col].ToString();
                                }
                            }
                            add = ExportProcess.AddColumn(add, 1);
                        }
                        iz++;
                    }
                }
            }
            #endregion
            return "OK";
        }
    }

}
