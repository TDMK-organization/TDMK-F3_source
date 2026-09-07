using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OfficeOpenXml;
using OK2SHIP_SMT.Repositories;

namespace OK2SHIP_SMT.Services
{
    public class ERS_NetSpecService
    {
        public Dictionary<string, DataTable> _DATA = new Dictionary<string, DataTable>();

        public ERS_NetSpecService()
        {
            _DATA.Add("Category", getDataTableStructor());
            _DATA["Category"].Rows.Add(" ");
        }

        public DataTable getDataTableStructor()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Pin1");
            dt.Columns.Add("Pin2");
            dt.Columns.Add("Net Name");
            dt.Columns.Add("Low Limit");
            dt.Columns.Add("High Limit");
            dt.Columns.Add("Min DCR");
            dt.Columns.Add("Max DCR");
            DataColumn colSelect = new DataColumn("Select", typeof(bool));
            colSelect.DefaultValue = false; // Giá trị mặc định là true (hoặc false tùy ý bạn)
            dt.Columns.Add(colSelect);
            return dt;
        }

        public void getData(string location, string itemCode, string maker)
        {
            _DATA.Clear();
            DataTable dt = getDataTableStructor();
            string FileName = FileFolderRepository.GetFileName(location);
            getInfo(location, out string itemCodeReal, out string makerReal);
            if (itemCodeReal != itemCode)
            {
                throw new Exception("ItemCode not match");
            }

            if (maker != makerReal)
            {
                throw new Exception("Maker not match");
            }

            ExportProcess _process = new ExportProcess();
            using (ExcelPackage package = ExportProcess.openPackage(location))
            {
                using (ExcelWorksheet worksheet = _process.FindSheet(package, "ERS_NET_SPEC"))
                {
                    IDictionary<string, string> dic = ExportProcess.FindAddressByText(worksheet,
                        new[] { "Pin1", "Pin2", "Net name", "LowLimit", "HighLimit", "Min DCR", "Max DCR", "ERS" });
                    int i = 1;
                    int colPin1 = worksheet.Cells[dic["Pin1"]].Start.Column;
                    int colPin2 = worksheet.Cells[dic["Pin2"]].Start.Column;
                    int colNetname = worksheet.Cells[dic["Net name"]].Start.Column;
                    int colLowLimit = worksheet.Cells[dic["LowLimit"]].Start.Column;
                    int colHighLimit = worksheet.Cells[dic["HighLimit"]].Start.Column;
                    int colMinDCR = worksheet.Cells[dic["Min DCR"]].Start.Column;
                    int colMaxDCR = worksheet.Cells[dic["Max DCR"]].Start.Column;
                    int colERS = worksheet.Cells[dic["ERS"].Split('-')[dic["ERS"].Split('-').Length - 1]].Start.Column;
                    int row = worksheet.Cells[dic["Max DCR"]].Start.Row;
                    while (true)
                    {
                        string valueNetName;
                        try
                        {
                            valueNetName = worksheet.Cells[row + i, colNetname].Value.ToString();
                            if (string.IsNullOrEmpty(valueNetName))
                            {
                                break;
                            }
                        }
                        catch
                        {
                            break;
                        }

                        DataRow rowValue = dt.NewRow();
                        rowValue["Net Name"] = valueNetName;
                        rowValue["Pin1"] = worksheet.Cells[row + i, colPin1].Value.ToString();
                        rowValue["Pin2"] = worksheet.Cells[row + i, colPin2].Value.ToString();
                        rowValue["Low Limit"] = worksheet.Cells[row + i, colLowLimit].Value.ToString();
                        rowValue["High Limit"] = worksheet.Cells[row + i, colHighLimit].Value.ToString();
                        rowValue["Min DCR"] = worksheet.Cells[row + i, colMinDCR].Value.ToString();
                        rowValue["Max DCR"] = worksheet.Cells[row + i, colMaxDCR].Value.ToString();
                        try
                        {
                            rowValue["Select"] =
                                worksheet.Cells[row + i, colERS].Value.ToString().ToUpper() == "YES" ? true : false;
                        }
                        catch
                        {
                            rowValue["Select"] = false;
                        }
                            dt.Rows.Add(rowValue);

                        i++;
                    }
                }
            }

            string[] arr = new[]
            {
                "Flex Bending", "Heat soak and Flex bend", "Thermal cycling and Flex bending", "Heat soak",
                "Thermal cycling", "Thermal shock"
            };
            foreach (string str in arr)
            {
                _DATA.Add(str, dt.Copy());
            }
        }

        public void getInfo(string location, out string itemCode, out string maker)
        {
            string FileName = FileFolderRepository.GetFileName(location);
            string[] split = FileName.Split('-');
            itemCode = split[0].Trim();
            maker = split[1].Trim();
        }

        private string _NAMETABLE = "BENDING_SPEC";
        private DBContext _DBCONTEXT = new DBContext();

        public void Save(string itemCode, string maker, bool prime = false)
        {
            DataTable dt = new DataTable();
            if (!prime)
            {
                dt =
                    _DBCONTEXT.LoadDataTable(_NAMETABLE, new[] { "ItemCode", "Maker" }, new[] { itemCode, maker });
                if (dt.Rows.Count > 0)
                {
                    throw new DataException("Đã có dữ liệu bạn có muốn tiếp tục?");
                }
            }
            else
            {
                dt = _DBCONTEXT.GetTableStructure(_NAMETABLE);
            }

            DataRow dr = dt.NewRow();
            dr["ItemCode"] = itemCode;
            dr["Maker"] = maker;

            List<string> data = new List<string>();
            foreach (string key in _DATA.Keys)
            {
                string value = $"{key}\u2060{ConverterService.DataTableToJson(_DATA[key])}";
                data.Add(value);
            }

            dr["Data"] = string.Join("\u200F", data);
            dt.Rows.Add(dr);

            _DBCONTEXT.BuckDataTable(dt, _NAMETABLE, new[] { "ItemCode", "Maker" }, null, "ID");
        }

        public void LoadData(string itemCode, string maker)
        {
            DataTable dt = _DBCONTEXT.LoadDataTable(_NAMETABLE, new[] { "ItemCode", "Maker" },
                new[] { itemCode.PadRight(20, ' '), maker.PadRight(20, ' ') });
            if (dt.Rows.Count <= 0)
            {
                dt = getDataTableStructor();
                dt.Rows.Add(" ");
                throw new Exception("Không có dữ liệu!");
            }

            DataRow row = dt.Rows[0];
            _DATA = new Dictionary<string, DataTable>();
            string data = row["Data"].ToString();
            string[] splits = data.Split('\u200F');
            foreach (string str in splits)
            {
                string[] zArr = str.Split('\u2060');
                _DATA.Add(zArr[0], ConverterService.JsonToDataTable(zArr[1]));
            }
        }

        public void ChangeValue(int col, string clipboardText, string category)
        {
            col = col < 0 ? 0 : col;
            Debugger.Break();
            string pattern = @"\b(yes|no|true|false)\b";

            MatchCollection matches = Regex.Matches(clipboardText, pattern, RegexOptions.IgnoreCase);

            // 3. Chuyển kết quả Regex thành danh sách bool
            List<bool> boolList = new List<bool>();
            foreach (Match match in matches)
            {
                string value = match.Value.ToLower(); // Đưa về chữ thường để dễ so sánh

                if (value == "yes" || value == "true")
                {
                    boolList.Add(true);
                }
                else if (value == "no" || value == "false")
                {
                    boolList.Add(false);
                }
            }

            // Lấy ra mảng bool[] như bạn yêu cầu
            bool[] boolArray = boolList.ToArray();


            // 4. Thay giá trị vào DataTable
            // Lưu ý: Cần phòng trường hợp số lượng giá trị trong clipboard khác số dòng của DataTable
            int rowCountToUpdate = boolArray.Length + col - 1 > _DATA[category].Rows.Count
                ? _DATA[category].Rows.Count
                : boolArray.Length;

            for (int i = 0; i < rowCountToUpdate; i++)
            {
                _DATA[category].Rows[i + col - 1]["Select"] = boolArray[i];
            }
        }
    }
}