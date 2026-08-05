using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using NationalInstruments.Restricted;
using OK2SHIP_SMT.Repositories;

namespace OK2SHIP_SMT.Services
{
    public class BendingService
    {
        private DBContext _dbContext = new DBContext();

        public static void LocationHandle(string location, out string itemCode, out string lotNo, out string maker)
        {
            location = location.Trim();
            itemCode = "";
            lotNo = "";
            maker = "";
            if (string.IsNullOrEmpty(location))
            {
                return;
            }

            location = FileFolderRepository.GetFolderName(location);
            try
            {
                string[] split = location.Split('-');
                string[] splitIL = split[0].Trim().Replace("00000B", "_").Split('_');
                itemCode = splitIL[0];
                lotNo = splitIL[1];
                maker = split[split.Length - 1];
            }
            catch
            {
                Debugger.Break();
            }
        }

        public Dictionary<string, DataTable> _RESULT = new Dictionary<string, DataTable>();

        public void GetLogfile(string location, string itemCode, string lotNo, string maker)
        {
            LocationHandle(location, out string itemCodes, out string lotNos, out string makers);
            if (!itemCode.Equals(itemCodes) || !lotNo.Equals(lotNos) || !maker.Equals(makers))
            {
                throw new Exception("Itemcode, lotno, makers is not match!");
            }

            string[] subs = FileFolderRepository.GetSubFolders(location);

            foreach (string folderAdd in subs)
            {
                string nameFolder = FileFolderRepository.GetFolderName(folderAdd).Trim();
                string key = "";

                switch (nameFolder)
                {
                    case "TC":
                        key = "Thermal cycling";
                        break;
                    case "F":
                        key = "Flex Bending";
                        break;
                    case "HF":
                        key = "Heat soak and Flex bend";
                        break;
                    case "TF":
                        key = "Thermal cycling and Flex bending";
                        break;
                    case "TS":
                        key = "Thermal shock";
                        break;
                    case "HS":
                        key = "Heat soak";
                        break;
                    default:
                        key = "EXCEPTION";
                        break;
                }

                DataTable dt = SolveCategory(folderAdd, nameFolder);
                if (_RESULT.TryGetValue(key, out _))
                {
                    // Tại sao lại có key giống nhau ở đây???
                    Debugger.Break();
                }
                else
                {
                    _RESULT.Add(key, dt);
                }
            }
        }

        private Dictionary<string, string> _DATA = new Dictionary<string, string>();

        private DataTable SolveCategory(string location, string category)
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Cycle");

            string[] cycles = FileFolderRepository.GetSubFolders(location).OrderBy(folderPath =>
            {
                // Lấy tên thư mục ở cuối (đề phòng mảng trả về đường dẫn tuyệt đối kiểu C:\...\L1)
                string folderName = Path.GetFileName(folderPath);

                // 1. Ưu tiên "BF" lên đầu tiên bằng cách gán cho nó giá trị sắp xếp nhỏ nhất
                if (folderName.Equals("BF", StringComparison.OrdinalIgnoreCase))
                {
                    return -1;
                }

                // 2. Nếu là thư mục "L", tách phần số phía sau để sắp xếp
                if (folderName.StartsWith("L", StringComparison.OrdinalIgnoreCase))
                {
                    // Chuyển đổi phần chuỗi sau chữ "L" thành số nguyên (VD: "L10" -> cắt lấy "10" -> số 10)
                    if (int.TryParse(folderName.Substring(1), out int number))
                    {
                        return number;
                    }
                }

                // 3. Đẩy bất kỳ thư mục nào không đúng chuẩn (nếu có) xuống cuối danh sách
                return int.MaxValue;
            }).ToArray();
            ;

            foreach (string cycle in cycles)
            {
                string cycleName = FileFolderRepository.GetFolderName(cycle);
                string[] getFiles = FileFolderRepository.GetFileByExtension(cycle, ".DAT");
                DataRow row = dataTable.NewRow();
                row["Cycle"] = cycleName;

                foreach (string file in getFiles)
                {
                    string fileName = FileFolderRepository.GetFileName(file).Trim();
                    string flexSN = fileName.Split('_')[3].Replace(".DAT", "");
                    if (!dataTable.Columns.Contains(flexSN))
                    {
                        dataTable.Columns.Add(flexSN);
                    }

                    row[flexSN] = HandleFile(file, flexSN, cycleName, category);
                }

                dataTable.Rows.Add(row);
            }

            return dataTable;
        }

        private string HandleFile(string location, string flexSN, string cycleName, string category)
        {
            string result = "N/A";
            string[] lines = File.ReadAllLines(location);
            string[] row = lines[0].Split(',');
            int measureI = row.IndexOf("\"Measure\""),
                tessResult = row.IndexOf("\"Test Title\""),
                lowI = row.IndexOf("\"Low-Lim\""),
                highI = row.IndexOf("\"High-Lim\"");
            List<string> listData = new List<string>();
            List<string> BFValue = new List<string>();
            if (category.Contains("F") && !cycleName.Contains("BF"))
            {
                if (_DATA.TryGetValue($"{flexSN}-{cycleName}-{category}", out string value))
                {
                    BFValue = value.Split('_').ToList();
                }
            }

            try
            {
                foreach (string line in lines.Skip(3))
                {
                    /// Điều kiện dừng khi 19M
                    /// 
                    if (result == "N/A")
                    {
                        result = "OK";
                    }

                    string[] data = line.Split(',');
                    string meanSureS = data[measureI].Replace("\"", "").Trim();
                    string lowS = data[lowI].Replace("\"", "").Trim();
                    string highS = data[highI].Replace("\"", "").Trim();
                    string tessResultS = data[tessResult].Replace("\"", "").Trim();


                    // check in range
                    if (double.TryParse(lowS, out double low) && double.TryParse(highS, out double high) &&
                        double.TryParse(meanSureS, out double measure))
                    {
                        if (low > measure || measure > high)
                        {
                            result = "NG";
                        }

                        if (BFValue.Count > 0)
                        {
                            string bf = BFValue[listData.Count - 1];
                            string meaSureBF = bf.Split('-')[0];
                            if (double.TryParse(meaSureBF, out double BFmeaSure))
                            {
                                double percent = Math.Abs((measure - BFmeaSure)) / BFmeaSure;
                                if ((percent > 10) || (percent < -10))
                                {
                                    result = "NG";
                                }
                            }
                        }

                        string s = $"{meanSureS}\u200B{tessResultS}\u200B{lowS}\u200B{highS}";
                        listData.Add(s);
                    }
                    else
                    {
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
            }

            _DATA.Add($"{flexSN}-{cycleName}-{category}", string.Join("\u200D", listData));

            return result;
        }

        private string switchCategory(string category)
        {
            switch (category)
            {
                case "Thermal cycling":
                    return "TC";
                case "Flex Bending":
                    return "F";
                case "Heat soak and Flex bend":
                    return "HF";
                case "Thermal cycling and Flex bending":
                    return "TF";
                case "Thermal shock":
                    return "TS";
                case "Heat soak":
                    return "HS";

                case "TC":
                    return "Thermal cycling";
                case "F":
                    return "Flex Bending";
                case "HF":
                    return "Heat soak and Flex bend";
                case "TF":
                    return "Thermal cycling and Flex bending";
                case "TS":
                    return "Thermal shock";
                case "HS":
                    return "Heat soak";
                default:
                    return "EXCEPTION";
            }
        }

        public DataTable GetDetail(string category, string flexSn)
        {
            DataTable dt = new DataTable();
            List<string> cycles = _RESULT[category].AsEnumerable()
                .Select(row => row[0]?.ToString() ?? "")
                .ToList();
            dt.Columns.Add("Net No", typeof(string));
            dt.Columns.Add("Flex SN", typeof(string));
            dt.Columns.Add("Test Result", typeof(string));
            dt.Columns.Add("LSL", typeof(string));
            dt.Columns.Add("USL", typeof(string));

            foreach (string cycle in cycles)
            {
                if (_DATA.TryGetValue($"{flexSn}-{cycle}-{switchCategory(category)}", out string value))
                {
                    dt.Columns.Add(cycle);
                    string[] split = value.Split('\u200D');
                    while (split.Length > dt.Rows.Count)
                    {
                        DataRow row = dt.NewRow();
                        row["Net No"] = dt.Rows.Count + 1;
                        row["Flex SN"] = flexSn;
                        row["Test Result"] = split[dt.Rows.Count].Split('\u200B')[1];
                        ;
                        dt.Rows.Add(row);
                    }

                    for (int si = 0; si < split.Length; si++)
                    {
                        string z = split[si].Split('\u200B')[3];
                        dt.Rows[si][cycle] = split[si].Split('\u200B')[0];
                        dt.Rows[si]["LSL"] = split[si].Split('\u200B')[2];
                        dt.Rows[si]["USL"] = split[si].Split('\u200B')[3];
                    }
                }
            }

            return dt;
        }

        private string _NAMETABLE = "BENDING_TEST_DATA";

        public void
            Save(string itemCode, string lotNo, string maker, int skip) // 0: chưa kiểm tra, 1: ghi đè, số khác: ghi mới
        {
            DataTable dt;
            if (skip == 0 || skip == 1)
            {
                dt = _dbContext.LoadDataTable(_NAMETABLE, new[] { "ItemCode", "LotNo", "Maker" },
                    new[] { itemCode, lotNo, maker });
                if (dt.Rows.Count > 0 && skip == 0)
                {
                    throw new DataException("Data is valid continue?");
                }

                if (skip == 1)
                {
                    //Kết hợp data từ database ra
                    List<string> result = dt.Rows[0]["Result"].ToString().Split('\u2060').ToList();
                    List<string> dataZ = dt.Rows[0]["Data"].ToString().Split('\u2060').ToList();
                    foreach (string s in dataZ)
                    {
                        string[] split = s.Split('\u200F');
                        if (!_DATA.TryGetValue(split[0], out string _))
                        {
                            _DATA.Add(split[0], split[1]);
                        }
                    }
                    // tính lại kết quả 

                    foreach (string s in result)
                    {
                        string[] split = s.Split('\u200F');
                        if (_RESULT.TryGetValue(split[0], out DataTable dtz))
                        {
                            _RESULT [split[0]]= SolveResult(split[0]);
                        }
                        else
                        {
                            _RESULT.Add(split[0], ConverterService.JsonToDataTable(split[1]));
                        }
                    }
                }
            }
            else
            {
                dt = _dbContext.GetTableStructure(_NAMETABLE);
            }

            List<string> data = new List<string>();
            foreach (string key in _DATA.Keys)
            {
                data.Add($"{key}\u200F{_DATA[key]}");
            }

            string dataJson = string.Join("\u2060", data);

            List<string> _result = new List<string>();
            foreach (string key in _RESULT.Keys)
            {
                _result.Add($"{key}\u200F{ConverterService.DataTableToJson(_RESULT[key])}");
            }

            string resultJson = string.Join("\u2060", _result);

            DataRow row;
            if (skip == 1)
            {
                row = dt.Rows[0];
            }
            else
            {
                row = dt.NewRow();
            }

            row["ItemCode"] = itemCode;
            row["LotNo"] = lotNo;
            row["Maker"] = maker;
            row["Data"] = dataJson;
            row["Result"] = resultJson;
            if (skip != 1)
            {
                dt.Rows.Add(row);
            }

            _dbContext.BuckDataTable(dt, _NAMETABLE, new[] { "ItemCode", "LotNo", "Maker" }, null, "ID");
        }

        /// <summary>
        /// This function is solve again result of category
        /// </summary>
        /// <param name="category"></param>
        /// <exception cref="NotImplementedException"></exception>
        private DataTable SolveResult(string category)
        {
            category = switchCategory(category);
            DataTable dt = new DataTable();
            dt.Columns.Add("Cycle", typeof(string));

            // tính dữ liệu BF trước và các cột
            // key {flexSN}-{cycleName}-{category}
            // Cycle - List<key>
            Dictionary<string, List<string>> _dic = new Dictionary<string, List<string>>();
            foreach (string key in _DATA.Keys)
            {
                string[] keySplit = key.Split('-');
                if (keySplit[2] == category)
                {
                    if (_dic.TryGetValue(keySplit[1], out List<string> list))
                    {
                        list.Add(key);
                        //_dic[keySplit[1]] = list;
                    }
                    else
                    {
                        _dic.Add(keySplit[1], new[] { key }.ToList());
                    }
                }
            }

            foreach (string s in _dic[_dic.Keys.First()])
            {
                dt.Columns.Add(s.Split('-')[0]);
            }

            List<string> keys = _dic.Keys
                .ToList() // 1. Đưa "BF" lên đầu tiên (nếu là "BF" trả về 0, còn lại trả về 1 để "BF" đứng trước)
                .OrderBy(k => k == "BF" ? 0 : 1)

                // 2. Sắp xếp các phần tử còn lại theo số đứng sau chữ "L"
                .ThenBy(k =>
                {
                    // Kiểm tra xem chuỗi có bắt đầu bằng "L" không và thử ép kiểu phần còn lại thành số nguyên
                    if (k.StartsWith("L") && int.TryParse(k.Substring(1), out int num))
                    {
                        return num;
                    }

                    // Nếu có key lạ không đúng định dạng, đẩy nó xuống cuối cùng
                    return int.MaxValue;
                })
                .ToList();
           // Debugger.Break();
            foreach (string cycle in keys)
            {
                DataRow row = dt.NewRow();
                row["Cycle"] = cycle;
                //{flexSN}-{cycleName}-{category}
                foreach (DataColumn col in dt.Columns)
                {
                    string flexSN = col.ColumnName;
                    if (_DATA.TryGetValue($"{flexSN}-{cycle}-{category}", out string data))
                    {
                        string[] listData = data.Split('\u200D');
                        string res = "N/A";
                        foreach (string s in listData)
                        {
                            res = "OK";
                            string[] splitS = s.Split('\u200B');
                            if (double.TryParse(splitS[2], out double low) &&
                                double.TryParse(splitS[3], out double high) &&
                                double.TryParse(splitS[0], out double measure))
                            {
                                if (measure < low || measure > high)
                                {
                                    res = "NG";
                                    break;
                                }

                                // So sánh với before không quá 10%
                                // if (dt.Rows.Count > 0 && category.Contains("F"))
                                // {
                                //     if (_DATA.TryGetValue($"{flexSN}-BF-{category}", out string z))
                                //     {
                                //         string meansure = z.Split('\u200D')[0];
                                //     }
                                // }
                            }
                            else
                            {
                                res = "N/A";
                                break;
                            }
                        }

                        row[flexSN] = res;
                    }
                }

                dt.Rows.Add(row);
            }

            return dt;
        }

        public void Load(string itemCode, string lotNo, string maker)
        {
            DataTable dt = _dbContext.LoadDataTable(_NAMETABLE, new[] { "ItemCode", "LotNo", "Maker" },
                new[] { itemCode, lotNo, maker });
            if (dt.Rows.Count <= 0)
            {
                throw new DataException("Data is empty");
            }

            List<string> result = dt.Rows[0]["Result"].ToString().Split('\u2060').ToList();
            List<string> data = dt.Rows[0]["Data"].ToString().Split('\u2060').ToList();
            foreach (string s in data)
            {
                string[] split = s.Split('\u200F');
                _DATA.Add(split[0], split[1]);
            }

            foreach (string s in result)
            {
                string[] split = s.Split('\u200F');
                _RESULT.Add(split[0], ConverterService.JsonToDataTable(split[1]));
            }
        }

        public void Export(string itemCode, string lotNo, string maker)
        {
        }

        public DataTable GetDetailByCycle(string category, string cycle)
        {
            cycle = cycle.Trim();
            DataTable dt = new DataTable();
            dt.Columns.Add("NetNo", typeof(string));
            Dictionary<string, string> FlexSNList = new Dictionary<string, string>();

            foreach (string key in _DATA.Keys)
            {
                //key : {flexSn}-{cycle}-{switchCategory(category)}
                string[] keySplit = key.Split('-');
                if (keySplit[1] == cycle && keySplit[2] == switchCategory(category))
                {
                    if (!FlexSNList.TryGetValue(keySplit[0], out string value))
                    {
                        string FlexSNKEY = keySplit[0];
                        string data = _DATA[$"{keySplit[0]}-{cycle}-{switchCategory(category)}"];
                        FlexSNList.Add(keySplit[0], _DATA[$"{keySplit[0]}-{cycle}-{switchCategory(category)}"]);
                        dt.Columns.Add($"ItemNo{FlexSNList.Count}", typeof(string));
                    }
                }
            }

            List<DataRow> listRow = new List<DataRow>();
            DataRow row = dt.NewRow();
            for (int i = 0; i < FlexSNList.Keys.Count; i++)
            {
                row[i + 1] = FlexSNList.Keys.ToList()[i];
                string[] split = FlexSNList[FlexSNList.Keys.ToList()[i]].Split('\u200D');
                while (split.Length > listRow.Count)
                {
                    DataRow rowZ = dt.NewRow();
                    rowZ["NetNo"] = listRow.Count + 1;
                    listRow.Add(rowZ);
                }

                for (int j = 0; j < split.Length; j++)
                {
                    listRow[j][i + 1] = split[j].Split('\u200B')[0];
                }
            }


            dt.Rows.Add(row);
            foreach (DataRow dataRow in listRow)
            {
                dt.Rows.Add(dataRow);
            }


            return dt;
        }

        public void UpdateData(DataTable dt, string category)
        {
            List<DataRow> listRow = new List<DataRow>();
            foreach (DataRow row in dt.Rows)
            {
                if ((bool)row["Selected"] == false)
                {
                    listRow.Add(row);
                }
            }

            foreach (DataRow row in listRow)
            {
                dt.Rows.Remove(row);
            }

            DataTable da = dt.Copy();
            da.Columns.RemoveAt(0);
            _RESULT[category] = da;
            Debugger.Break();
        }
    }
}