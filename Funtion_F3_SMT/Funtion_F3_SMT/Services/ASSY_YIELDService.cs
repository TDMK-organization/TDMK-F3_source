using OfficeOpenXml;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Security.Authentication;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.Services
{
    public class ASSY_YIELDService
    {
        private string _nameTable = "ASSY_YIELD";
        private DBContext _dbContext = new DBContext();
        public int Save(string itemCode, string lotNo, Dictionary<string, DataTable> dIC, Guid area)
        {
            int res = 0;
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            if (string.IsNullOrEmpty(itemCode) && string.IsNullOrEmpty(lotNo))
            {
                throw new Exception("Không được để trống itemCode lotno");
            }
            if (!UserSession.Instance.IsLoggedIn)
            {
                throw new AuthenticationException("Hãy đăng nhập");
            }
            DataTable dataTable = new DataTable();
            // nếu chưa kiểm tra database thì -1
            if (area == Guid.Parse("00000000-0000-0000-0000-000000000000"))
            {
                dataTable = _dbContext.LoadDataTable(_nameTable, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
                if (dataTable.Rows.Count > 0)
                {
                    throw new Exception($"Area = {dataTable.Rows[0]["Area"]} & Đã có dữ liệu bạn có muốn ghi đè!");
                }
                area = Guid.NewGuid();
            }
            // nếu đã kiểm tra trả về area
            else
            {
                dataTable = _dbContext.GetTableStructure(_nameTable);
            }
            DataRow row = dataTable.NewRow();
            DataTable imageDataTable = _dbContext.GetTableStructure(_nameTable + "_IMAGE");
            int id = 1;
            foreach (string item in dIC.Keys)
            {
                switch (item)
                {
                    case "Process":
                        row[item] = ConverterService.DataTableToJson(dIC[item]);
                        break;
                    case "Top_SMT":
                    case "Top_Backend":
                    case "OQC":
                        string json = ConverterService.ConvertDataTableImage(dIC[item], imageDataTable, ref id, area);
                        row[item] = json;
                        break;
                    default:
                        break;

                }
            }
            row["ItemCode"] = itemCode;
            row["LotNo"] = lotNo;
            row["Area"] = area;

            dataTable.Rows.Add(row);
            res += _dbContext.BuckDataTable(imageDataTable, _nameTable + "_IMAGE", new[] { "Area" }, null, "ID");
            res += _dbContext.BuckDataTable(dataTable, _nameTable, new[] { "ItemCode", "LotNo" }, null, "ID");
            return res;
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

                DataTable zz = ConverterService.JsonToDataTable(json);
                if (zz.Rows.Count <= 0)
                {
                    continue;
                }
                else
                {

                    DataTable z = ConverterService.ConvertDataTableImage(zz, dataImage);
                    dIC.Add(item, z);
                }

            }
            return dIC;

        }

        public void Export(string itemCode, string lotNo)
        {
            if (UserSession.Instance.IsLoggedIn == false)
            {
                throw new AuthenticationException("Hãy đăng nhập");
            }

            Dictionary<string, DataTable> dIC = Load(itemCode, lotNo);
            ExportProcess exportProcess = new ExportProcess();
            using (ExcelPackage package = exportProcess.FindFormatProcess("Assy Yield", itemCode, lotNo))
            {
                using (ExcelWorksheet worksheet = exportProcess.FindSheet(package, "Assy Yield"))
                {
                    string[] healder = new string[] { "Station", "Input", "Passed and shipped to next process", "Rejected", "Evaluation", "IPQC", "ORT", "WIP", "Others" };
                    IDictionary<string, string> dic = ExportProcess.FindAddressByText(worksheet, healder, true);
                    #region Process
                    DataTable dataTable = dIC["Process"];
                    Dictionary<string, int> Marking = new Dictionary<string, int>();
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        DataRow row = dataTable.Rows[i];
                        Marking.Add(row["Station"].ToString(), i);
                    }
                    string address = dic["Station"].Split('-')[dic["Station"].Split('-').Count() - 1];
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
                                        if (row[col] is byte[])
                                        {
                                            ExportProcess.InsertImageToCell(worksheet, newz, (byte[])row[col], $"{Guid.NewGuid()}");
                                        }
                                        else
                                        {
                                            ExportProcess.InsertImageToCell(worksheet, newz, TDMK_ImageConverter.ImageToByteArray((Image)row[col], ImageFormat.Png), $"{Guid.NewGuid()}");
                                        }

                                        add = ExportProcess.AddColumn(add, 1);

                                    }
                                    else
                                    {
                                        if (col.ColumnName.Equals("Defect Name"))
                                        {
                                            add = ExportProcess.AddColumn(add, -1);

                                        }
                                        worksheet.Cells[add].Value = row[col].ToString();
                                    }
                                    add = ExportProcess.AddColumn(add, 1);
                                }
                                iz++;
                            }
                        }
                    }
                    #endregion
                    exportProcess.SaveExcelWorksheet(package, "Assy Yield", $"Assy Yield - {itemCode} - {lotNo}");
                }
            }

        }
    }
}
