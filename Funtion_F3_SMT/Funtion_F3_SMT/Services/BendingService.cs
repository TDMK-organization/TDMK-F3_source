using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using NationalInstruments.Restricted;
using OfficeOpenXml;
using OK2SHIP_SMT.Repositories;

namespace OK2SHIP_SMT.Services
{
    public class BendingService
    {
        private DBContext _dbContext = new DBContext();
        private Dictionary<string, DataTable> _RESULT = new Dictionary<string, DataTable>();
        private DataTable __SPEC = new DataTable();
        private Dictionary<string, string> _DATA = new Dictionary<string, string>();

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


        public DataTable getValue(string key)
        {
            if (_RESULT.TryGetValue(key, out DataTable dt))
            {
                DataTable result = dt.Copy();
                if (!result.Columns.Contains("Selected"))
                {
                    DataColumn checkCol = new DataColumn("Selected", typeof(bool))
                    {
                        DefaultValue = true
                    };
                    result.Columns.Add(checkCol);
                    checkCol.SetOrdinal(0);
                }


                return result;
            }

            throw new DataException("Dữ liệu không tồn tại!");
        }

        public List<string> ListKey()
        {
            return _RESULT.Keys.ToList();
        }

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

        // Key         
        // "FlexSN"-"Cycle"-"Category"-"NetNo" = "Detail Error Item"    
        // Đếm số lượng NG của mỗi mục trên
        public Dictionary<string, string> _ERROR_LIST = new Dictionary<string, string>();

        public DataTable totalErr(string key, out int total, out int lsl, out int usl, out int rng, out int ok,
            out int tng)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Flex SN");
            dt.Columns.Add("Cycle");
            dt.Columns.Add("Net No");
            dt.Columns.Add("ERROR");
            total = _ERROR_LIST.Keys.Count;
            lsl = 0;
            usl = 0;
            rng = 0;
            ok = 0;
            tng = 0;

            string category = switchCategory(key);
            foreach (string keyItem in _ERROR_LIST.Keys)
            {
                string[] split = keyItem.Split('-');
                if (split[2] == category)
                {
                    string cycle = split[1];


                    string s = _ERROR_LIST[keyItem];
                    if (s != "OK")
                    {
                        string[] splitZZ = keyItem.Split('-');
                        DataRow row = dt.NewRow();
                        row["Flex SN"] = splitZZ[0];
                        row["Cycle"] = splitZZ[1];
                        row["Net No"] = splitZZ[3];
                        row["ERROR"] = s;
                        dt.Rows.Add(row);
                    }

                    switch (s)
                    {
                        case "Over USL":
                            usl++;
                            break;
                        case "Over LSL":
                            lsl++;
                            break;
                        case "Not Valid":
                            Debugger.Break();
                            break;
                        case "Vanability R NG":
                            rng++;
                            break;
                        case "Vanability T NG":
                            tng++;
                            break;
                        case "OK":
                            ok++;
                            break;
                        default:
                            Debugger.Break();
                            break;
                    }
                }
            }

            List<string> es = new List<string>();
            // foreach (string keyz in dic.Keys)
            // {
            //     es.Add($"{keyz}:{dic[keyz]}");
            // }

            return dt;
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
                if (_DATA.TryGetValue($"{flexSN}-BF-{category}", out string value))
                {
                    BFValue = value.Split('\u200D').ToList();
                }
            }

            try
            {
                foreach (string line in lines.Skip(3))
                {
                    string result_net = "Not Valid";

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
                            if (low > measure)
                            {
                                result_net = "Over USL";
                            }
                            else
                            {
                                result_net = "Over LSL";
                            }
                        }
                        else
                        {
                            result_net = "OK";
                        }

                        if (BFValue.Count > 0 && result_net == "OK")
                        {
                            string bf = BFValue[listData.Count];
                            string meaSureBF = bf.Split('\u200B')[0];
                            if (double.TryParse(meaSureBF, out double BFmeaSure))
                            {
                                double percent = (measure - BFmeaSure) / BFmeaSure;
                                if ((percent > 0.1) || (percent < -0.1))
                                {
                                    result = "NG";
                                    result_net = "Vanability R NG";
                                }
                                else
                                {
                                    if ((percent > 0.095) || (percent < -0.095))
                                    {
                                        result_net = "Vanability T NG";
                                    }
                                    else
                                    {
                                        result_net = "OK";
                                    }
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

                    _ERROR_LIST.Add($"{flexSN}-{cycleName}-{category}-{listData.Count}", result_net);
                }
            }
            catch (Exception ex)
            {
                Debugger.Break();
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

            if (__SPEC.Rows.Count > 0)
            {
                if (!dt.Columns.Contains("Net Name"))
                {
                    // 1. Thêm cột mới vào result (mặc định sẽ nằm ở cuối bảng)
                    DataColumn newColumn = dt.Columns.Add("Net Name", typeof(string));

                    // 2. Di chuyển cột này về vị trí thứ 2 (Index = 1)
                    newColumn.SetOrdinal(3);

                    // 3. Copy dữ liệu từ __SPEC sang result theo từng dòng
                    for (int i = 0; i < Math.Min(__SPEC.Rows.Count, dt.Rows.Count); i++)
                    {
                        if (__SPEC.Columns.Contains("Net Name"))
                        {
                            dt.Rows[i]["Net Name"] = __SPEC.Rows[i]["Net Name"];
                        }
                    }
                }

                if (!dt.Columns.Contains("Pin1"))
                {
                    // 1. Thêm cột mới vào result (mặc định sẽ nằm ở cuối bảng)
                    DataColumn newColumn = dt.Columns.Add("Pin1", typeof(string));

                    // 2. Di chuyển cột này về vị trí thứ 2 (Index = 1)
                    newColumn.SetOrdinal(4);

                    // 3. Copy dữ liệu từ __SPEC sang result theo từng dòng
                    for (int i = 0; i < Math.Min(__SPEC.Rows.Count, dt.Rows.Count); i++)
                    {
                        if (__SPEC.Columns.Contains("Pin1"))
                        {
                            dt.Rows[i]["Pin1"] = __SPEC.Rows[i]["Pin1"];
                        }
                    }
                }

                if (!dt.Columns.Contains("Pin2"))
                {
                    // 1. Thêm cột mới vào result (mặc định sẽ nằm ở cuối bảng)
                    DataColumn newColumn = dt.Columns.Add("Pin2", typeof(string));

                    // 2. Di chuyển cột này về vị trí thứ 2 (Index = 1)
                    newColumn.SetOrdinal(5);

                    // 3. Copy dữ liệu từ __SPEC sang result theo từng dòng
                    for (int i = 0; i < Math.Min(__SPEC.Rows.Count, dt.Rows.Count); i++)
                    {
                        if (__SPEC.Columns.Contains("Pin2"))
                        {
                            dt.Rows[i]["Pin2"] = __SPEC.Rows[i]["Pin2"];
                        }
                    }
                }

                if (!dt.Columns.Contains("Select"))
                {
                    // 1. Thêm cột mới vào result (mặc định sẽ nằm ở cuối bảng)
                    DataColumn newColumn = dt.Columns.Add("Select", typeof(string));

                    // 2. Di chuyển cột này về vị trí thứ 2 (Index = 1)
                    newColumn.SetOrdinal(5);

                    // 3. Copy dữ liệu từ __SPEC sang result theo từng dòng
                    for (int i = 0; i < Math.Min(__SPEC.Rows.Count, dt.Rows.Count); i++)
                    {
                        if (__SPEC.Columns.Contains("Select"))
                        {
                            dt.Rows[i]["Select"] = __SPEC.Rows[i]["Select"];
                        }
                    }
                }
            }

            return dt;
        }

        private string _NAMETABLE = "BENDING_TEST_DATA";

        public List<string> getListFlexSN(string itemName)
        {
            List<string> result = new List<string>();
            return result;
        }

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
                            _RESULT[split[0]] = SolveResult(split[0]);
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
            List<string> errorList = new List<string>();
            foreach (string key in _ERROR_LIST.Keys)
            {
                errorList.Add($"{key}\u200D{_ERROR_LIST[key]}");
            }

            string jsonErr = string.Join("\u200B", errorList);

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
            row["ErrorList"] = jsonErr;
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
                        // So sánh với before không quá 10%
                        string[] listBF = null;
                        if (dt.Rows.Count > 0 && category.Contains("F") && cycle != "BF")
                        {
                            if (_DATA.TryGetValue($"{flexSN}-BF-{category}", out string bf))
                            {
                                listBF = bf.Split('\u200D');
                            }
                        }

                        string res = "N/A";
                        string result_err = "Not Valid";
                        for (int i = 0; i < listData.Length; i++)
                        {
                            string s = listData[i];
                            if (res == "N/A")
                            {
                                res = "OK";
                            }

                            result_err = "OK";
                            string[] splitS = s.Split('\u200B');
                            if (double.TryParse(splitS[2], out double low) &&
                                double.TryParse(splitS[3], out double high) &&
                                double.TryParse(splitS[0], out double measure))
                            {
                                if (measure < low || measure > high)
                                {
                                    res = "NG";
                                    if (low > measure)
                                    {
                                        result_err = "Over USL";
                                    }
                                    else
                                    {
                                        result_err = "Over LSL";
                                    }
                                }

                                if (listBF != null && i < listBF.Length)
                                {
                                    string[] splitBF = listBF[i].Split('\u200B');
                                    if (double.TryParse(splitBF[0], out double valueBF))
                                    {
                                        double z = (valueBF - measure) / valueBF;
                                        // Debugger.Break();
                                        if (z > 0.1 || z < -0.1)
                                        {
                                            res = "NG";
                                            result_err = "Vanability R NG";
                                        }
                                        else if ((z > 0.095) || (z < -0.095))
                                        {
                                            result_err = "Vanability T NG";
                                        }
                                    }
                                }
                            }
                            else
                            {
                                res = "N/A";

                                result_err = "N/A";
                            }

                            if (_ERROR_LIST.TryGetValue($"{flexSN}-{cycle}-{category}-{i}", out string _))
                            {
                                _ERROR_LIST[$"{flexSN}-{cycle}-{category}-{i}"] = result_err;
                            }
                            else
                            {
                                _ERROR_LIST.Add($"{flexSN}-{cycle}-{category}-{i}", result_err);
                            }
                        }

                        row[flexSN] = res;
                    }
                }

                dt.Rows.Add(row);
            }

            return dt;
        }

        private void loadSpec(string itemCode, string maker)
        {
            ERS_NetSpecService _service = new ERS_NetSpecService();
            _service.LoadData(itemCode, maker);
            __SPEC = _service._DATA;
        }

        public void Load(string itemCode, string lotNo, string maker)
        {
            _DATA.Clear();
            _RESULT.Clear();
            _ERROR_LIST.Clear();
            DataTable dt = _dbContext.LoadDataTable(_NAMETABLE, new[] { "ItemCode", "LotNo", "Maker" },
                new[] { itemCode, lotNo, maker });
            try
            {
                loadSpec(itemCode, maker);
            }
            catch
            {
            }

            if (dt.Rows.Count <= 0)
            {
                throw new DataException("Data is empty");
            }

            List<string> result = dt.Rows[0]["Result"].ToString().Split('\u2060').ToList();
            List<string> data = dt.Rows[0]["Data"].ToString().Split('\u2060').ToList();
            List<string> jsonErr = dt.Rows[0]["ErrorList"].ToString().Split('\u200B').ToList();
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

            foreach (string s in jsonErr)
            {
                string[] split = s.Split('\u200D');
                _ERROR_LIST.Add(split[0], split[1]);
            }
        }

        public void Export(string itemCode, string lotNo, string maker, string category = "Flex Bending")
        {
            string namesheet = "", ictsheet = "", xraysheet = "";
            bool primeBend = false;
            switch (category)
            {
                case "Flex Bending":
                    namesheet = "Flex bending";
                    ictsheet = "Flex bending - ICT data";
                    xraysheet = "Flex Bending - X-Ray pictures";
                    primeBend = true;
                    break;
                case "Heat soak and Flex bend":
                    namesheet = "Heat Soak & bending";
                    ictsheet = "HS & bending - ICT data";
                    xraysheet = "HS & bending - X-Ray pictures";
                    primeBend = true;
                    break;
                case "Thermal cycling and Flex bending":
                    namesheet = "Thermal Cycling & bending";
                    ictsheet = "TC & bending - ICT data";
                    xraysheet = "TC & bending - X-Ray pictures";
                    primeBend = true;
                    break;
                case "Heat soak":
                    namesheet = "Heat Soak and Recovery";
                    ictsheet = "Heat soak & Recovery - ICT";
                    break;
                case "Thermal cycling":
                    namesheet = "Thermal Cycling";
                    ictsheet = "Thermal Cycling - ICT data";
                    break;
                case "Thermal shock":
                    namesheet = "Thermal Shock ";
                    ictsheet = "Thermal Shock - ICT data";
                    break;

                default:
                    break;
            }

            if (string.IsNullOrEmpty(maker))
            {
                DataTable dt_maker = _dbContext.LoadDataTable(_NAMETABLE, new[] { "ItemCode", "LotNo" },
                    new[] { itemCode, lotNo }, new[] { "ItemCode", "LotNo", "Maker" });
                ExportProcess process = new ExportProcess();

                using (ExcelPackage package = process.FindFormatProcess("", itemCode, lotNo))
                {
                    List<string> name = new[] { namesheet, ictsheet }.ToList();

                    using (ExcelWorksheet worksheet = process.FindSheet(package, namesheet))
                    using (ExcelWorksheet worksheet_ICT = process.FindSheet(package, ictsheet))
                    using (ExcelWorksheet worksheetXray = process.FindSheet(package, xraysheet))
                    {
                        int slot = 0;
                        if (dt_maker.Rows.Count > 0)
                        {
                            SetupResultSheet(worksheet, primeBend, dt_maker.Rows.Count);
                            if (primeBend)
                            {
                                SetupXray(worksheetXray, dt_maker.Rows.Count);
                            }
                        }

                        List<string> makers = new List<string>();
                        foreach (DataRow row in dt_maker.Rows)
                        {
                            itemCode = row["ItemCode"].ToString();
                            lotNo = row["LotNo"].ToString();
                            maker = row["Maker"].ToString();

                            Load(itemCode, lotNo, maker);

                            if (primeBend)
                            {
                                DataTable dt = _RESULT[category];
                                ExportDataToWS(worksheet, category, dt, maker, slot);

                                // Tránh add trùng tên xraysheet nhiều lần vào danh sách name
                                if (!name.Contains(xraysheet))
                                {
                                    name.Add(xraysheet);
                                }

                                makers.Add(maker);
                            }
                            else
                            {
                                ExportDataToWSNotBend(worksheet, category, slot);
                            }

                            // ==========================================
                            // XỬ LÝ DUPLICATE ICT SHEET CHO NHIỀU MAKER
                            // ==========================================
                            ExcelWorksheet currentIctSheet = worksheet_ICT; // Mặc định Maker 1 xài sheet gốc

                            if (slot > 0)
                            {
                                // Tạo tên sheet mới (VD: ICT1, ICT2...)
                                string newIctSheetName = $"{ictsheet}{slot}";

                                // Duplicate sheet từ sheet ICT gốc
                                currentIctSheet = package.Workbook.Worksheets.Add(newIctSheetName, worksheet_ICT);

                                // Đưa tên sheet mới vào danh sách giữ lại (tránh bị xóa ở bước Save)
                                name.Add(newIctSheetName);
                            }

                            // Đổ dữ liệu vào sheet tương ứng với Maker hiện tại
                            ExportDataToICT(currentIctSheet, category, primeBend);
                            // ==========================================

                            slot++;
                        }

                        try
                        {
                            if (makers.Count > 0) // Cẩn thận kiểm tra tránh truyền list rỗng
                            {
                                XRayPictureService service = new XRayPictureService();
                                service.ExportMultipleMakers(worksheetXray, itemCode, lotNo, makers, category);
                            }
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            Debugger.Break();
#endif
                        }

                        package.Compression = OfficeOpenXml.CompressionLevel.BestSpeed; // Đã ép xung tốc độ Save
                        process.SaveExcelWorksheet(package, string.Join(":", name), $"{itemCode}-{lotNo}-{category}");
                    }
                }
            }
            else
            {
                ExportProcess process = new ExportProcess();
                Load(itemCode, lotNo, maker);
                using (ExcelPackage package = process.FindFormatProcess("", itemCode, lotNo))
                {
                    List<string> name = new[] { namesheet, ictsheet }.ToList();
                    using (ExcelWorksheet worksheet = process.FindSheet(package, namesheet))
                    using (ExcelWorksheet worksheet_ICT = process.FindSheet(package, ictsheet))
                    using (ExcelWorksheet worksheetXray = process.FindSheet(package, xraysheet))
                    {
                        if (primeBend)
                        {
                            DataTable dt = _RESULT[category];
                            ExportDataToWS(worksheet, category, dt, maker);
                            XRayPictureService service = new XRayPictureService();
                            name.Add(xraysheet);
                            try
                            {
                                service.Export(worksheetXray, itemCode, lotNo, maker, category, false, 0);
                            }
                            catch
                            {
                                Debugger.Break();
                            }
                        }
                        else
                        {
                            ExportDataToWSNotBend(worksheet, category, 0);
                        }

                        ExportDataToICT(worksheet_ICT, category, primeBend);
                        package.Compression = OfficeOpenXml.CompressionLevel.BestSpeed;
                        process.SaveExcelWorksheet(package, string.Join(":", name), $"{itemCode}-{lotNo}");
                    }
                }
            }
        }

        private void SetupXray(ExcelWorksheet worksheet, int slot)
        {
            IDictionary<string, string> _dic = ExportProcess.FindAddressByText(worksheet, new[]
                { "Flex SN", "Result", "#" });
            ExportProcess process = new ExportProcess();
            string sample = _dic["#"].Split('-')[_dic["#"].Split('-').Length - 1];
            string check = _dic["Result"].Split('-')[_dic["Result"].Split('-').Length - 1];
            string rangeBase =
                $"{_dic["Flex SN"]}:{worksheet.Cells[worksheet.Cells[check].Start.Row, worksheet.Cells[sample].Start.Column].Address}";
            string newAdd = worksheet.Cells[ExportProcess.AddRow(check, 3)].Address;
            for (int i = 1; i < slot; i++)
            {
                process.CopyAndInsert(worksheet, rangeBase, ref newAdd);
            }
        }

        private void SetupResultSheet(ExcelWorksheet worksheet, bool bend, int slot)
        {
            if (bend)
            {
                IDictionary<string, string> _dic = ExportProcess.FindAddressByText(worksheet, new[]
                    { "Condition", "Cosmetic check", "Sample" });
                ExportProcess process = new ExportProcess();
                string sample = _dic["Sample"].Split('-')[_dic["Sample"].Split('-').Length - 1];
                string check = _dic["Cosmetic check"].Split('-')[_dic["Cosmetic check"].Split('-').Length - 1];
                string rangeBase =
                    $"{_dic["Condition"]}:{worksheet.Cells[worksheet.Cells[check].Start.Row, worksheet.Cells[sample].Start.Column].Address}";
                string newAdd = worksheet.Cells[ExportProcess.AddRow(check, 3)].Address;
                for (int i = 1; i < slot; i++)
                {
                    process.CopyAndInsert(worksheet, rangeBase, ref newAdd);
                }
            }
            else
            {
                IDictionary<string, string> _dic = ExportProcess.FindAddressByText(worksheet, new[]
                    { "Content", "Function test", "#" });
                ExportProcess process = new ExportProcess();
                string sample = _dic["#"].Split('-')[_dic["#"].Split('-').Length - 1];
                string check = _dic["Function test"].Split('-')[_dic["Function test"].Split('-').Length - 1];
                string rangeBase =
                    $"{_dic["Content"]}:{worksheet.Cells[worksheet.Cells[check].Start.Row, worksheet.Cells[sample].Start.Column].Address}";
                string newAdd = worksheet.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(check, 3), -1)].Address;
                for (int i = 1; i < slot; i++)
                {
                    process.CopyAndInsert(worksheet, rangeBase, ref newAdd);
                }
            }
        }

        private void ExportDataToWSNotBend(ExcelWorksheet ws, string category, int slot)
        {
            IDictionary<string, string> _dic =
                ExportProcess.FindAddressByText(ws, new[] { "Flex SN", "Sample no.", "Before" });
            _dic["Flex SN"] = _dic["Flex SN"].Split('-')[slot];
            _dic["Sample no."] = _dic["Sample no."].Split('-')[slot];
            _dic["Before"] = _dic["Before"].Split('-')[slot];
            DataTable dt = _RESULT[category];
            //Debugger.Break();
            string addressSample = ExportProcess.AddColumn(_dic["Sample no."], 2);
            string valueSample = ws.Cells[addressSample].Value.ToString();
            int rowFlexSN = ws.Cells[_dic["Flex SN"]].Start.Row;
            int i = 1;
            try
            {
                while (!string.IsNullOrEmpty(valueSample))
                {
                    ws.Cells[rowFlexSN, ws.Cells[addressSample].Start.Column].Value = dt.Columns[i].ColumnName;
                    try
                    {
                        string addresscycle = _dic["Before"];
                        string cycle = ws.Cells[addresscycle].Value.ToString();
                        while (!string.IsNullOrEmpty(cycle))
                        {
                            string cycleZ = "";
                            int num = GetFirstNumber(cycle);
                            if (0 == num)
                            {
                                cycleZ = "BF";
                            }
                            else
                            {
                                cycleZ = $"L{num / 100}";
                            }

                            string value = "N/A";

                            foreach (DataRow row in dt.Rows)
                            {
                                // Kiểm tra xem cột "Cycle" có khác null và giá trị có khớp với cycleZ không
                                if (row["Cycle"] != DBNull.Value && row["Cycle"].ToString() == cycleZ.ToString())
                                {
                                    // Lấy giá trị tại cột thứ i (hoặc tên cột)
                                    value = row[dt.Columns[i]].ToString();
                                    break; // Thoát vòng lặp ngay khi tìm thấy dòng đầu tiên thỏa mãn
                                }
                            }

                            ws.Cells[ws.Cells[addresscycle].Start.Row, ws.Cells[addressSample].Start.Column].Value =
                                value;

                            ws.Cells[ws.Cells[addresscycle].Start.Row + 1, ws.Cells[addressSample].Start.Column].Value =
                                value.Equals("OK") ? "Pass" : "Failed";

                            addresscycle = ExportProcess.AddRow(addresscycle, 2);
                            cycle = ws.Cells[addresscycle].Value.ToString();
                        }
                    }
                    catch
                    {
                    }

                    addressSample = ExportProcess.AddColumn(addressSample, 1);
                    valueSample = ws.Cells[addressSample].Value.ToString();
                }
            }
            catch
            {
                Debugger.Break();
            }
        }


        private void ExportDataToWS(ExcelWorksheet ws, string category, DataTable dt, string maker, int slot = 0)
        {
            IDictionary<string, string> _dic =
                ExportProcess.FindAddressByText(ws, new[] { "Flex SN", "E-check", "Sample", "Condition" });


            string[] splitSample = _dic["Sample"].Split('-');
            List<string> cache = new List<string>();
            string[] splitCondition = _dic["Condition"].Split('-');
            ws.Cells[ExportProcess.AddColumn(splitCondition[slot], 1)].Value =  maker;
            int rowS = ws.Cells[splitCondition[slot]].Start.Row;
            int rowE;
            try
            {
                rowE = ws.Cells[splitCondition[slot + 1]].Start.Row;
            }
            catch
            {
                rowE = -1;
            }

            foreach (string item in splitSample)
            {
                int r = ws.Cells[item].Start.Row;
                if (rowE == -1)
                {
                    if (r > rowS)
                    {
                        cache.Add(item);
                    }
                }
                else
                {
                    if (r > rowS && r < rowE)
                    {
                        cache.Add(item);
                    }
                }
            }

            splitSample = cache.ToArray();
            cache = new List<string>();
            string[] splitEcheck = _dic["E-check"].Split('-');
            foreach (string item in splitEcheck)
            {
                int r = ws.Cells[item].Start.Row;
                if (rowE == -1)
                {
                    if (r > rowS)
                    {
                        cache.Add(item);
                    }
                }
                else
                {
                    if (r > rowS && r < rowE)
                    {
                        cache.Add(item);
                    }
                }
            }

            splitEcheck = cache.ToArray();
            _dic["Flex SN"] = _dic["Flex SN"].Split('-')[slot];
            foreach (string sampleAdd in splitSample)
            {
                try
                {
                    string textSample = ws.Cells[sampleAdd].Text.ToString().Replace("Sample", "").Trim();
                    if (int.TryParse(textSample, out int sampleNo))
                    {
                        int column = ws.Cells[sampleAdd].Start.Column;
                        ws.Cells[ws.Cells[_dic["Flex SN"]].Start.Row, column].Value = dt.Columns[sampleNo].ColumnName;
                        foreach (string echeckAdd in splitEcheck)
                        {
                            if (!ws.Cells[echeckAdd].Text.Contains("Result"))
                            {
                                int cycle = GetFirstNumber(ws.Cells[echeckAdd].Text);

                                int row = ws.Cells[echeckAdd].Start.Row;
                                //Debugger.Break();
                                string valueCycle = "BF";
                                if (cycle != 0)
                                {
                                    valueCycle = $"L{cycle}";
                                }

                                ws.Cells[row, column].Value = "N/A";
                                ws.Cells[row + 1, column].Value = "N/A";
                                foreach (DataRow rowData in dt.Rows)
                                {
                                    if (rowData["Cycle"].ToString().Equals(valueCycle))
                                    {
                                        string value = rowData[dt.Columns[sampleNo].ColumnName].ToString();
                                        ws.Cells[row, column].Value = value;
                                        if (category == "Heat soak and Flex bend" && valueCycle == "BF")
                                        {
                                            Debugger.Break();
                                        }
                                        else
                                        {
                                            ws.Cells[row + 1, column].Value = value.Equals("OK") ? "Pass" : "Failed";
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debugger.Break();
                }
            }
        }

        private void ExportDataToICT(ExcelWorksheet ws, string category, bool primeBend = false)
        {
            // 1. Tìm các tọa độ gốc
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(ws,
                new[]
                {
                    "Sample no.", "Net no.", "Flex SN", "Sumitomo net name", "Customer net name", "Pin 1", "Pin 2",
                    "LSL", "USL", "Before"
                });

            dic["Before"] = dic["Before"].Split('-').Last();

            // 2. Chuyển đổi TẤT CẢ tọa độ chuỗi sang số nguyên (Row, Col) ĐÚNG 1 LẦN
            int row_Sample = ws.Cells[dic["Sample no."]].Start.Row;
            int col_Sample = ws.Cells[dic["Sample no."]].Start.Column;

            int row_Net = ws.Cells[dic["Net no."]].Start.Row;
            int col_Net = ws.Cells[dic["Net no."]].Start.Column;

            int row_Flex = ws.Cells[dic["Flex SN"]].Start.Row;
            int col_Flex = ws.Cells[dic["Flex SN"]].Start.Column;

            int row_Sumi = ws.Cells[dic["Sumitomo net name"]].Start.Row;
            int col_Sumi = ws.Cells[dic["Sumitomo net name"]].Start.Column;

            int row_Cus = ws.Cells[dic["Customer net name"]].Start.Row;
            int col_Cus = ws.Cells[dic["Customer net name"]].Start.Column;

            int row_Pin1 = ws.Cells[dic["Pin 1"]].Start.Row;
            int col_Pin1 = ws.Cells[dic["Pin 1"]].Start.Column;

            int row_Pin2 = ws.Cells[dic["Pin 2"]].Start.Row;
            int col_Pin2 = ws.Cells[dic["Pin 2"]].Start.Column;

            int row_LSL = ws.Cells[dic["LSL"]].Start.Row;
            int col_LSL = ws.Cells[dic["LSL"]].Start.Column;

            int row_USL = ws.Cells[dic["USL"]].Start.Row;
            int col_USL = ws.Cells[dic["USL"]].Start.Column;

            int row_Before = ws.Cells[dic["Before"]].Start.Row;
            int col_Before = ws.Cells[dic["Before"]].Start.Column;

            // 3. TIỀN XỬ LÝ CÁC CỘT ĐỘNG (CYCLE): CHỈ QUÉT ĐÚNG 1 LẦN
            Dictionary<int, string> cycleColumns = new Dictionary<int, string>();
            int headerRow = row_Before; // Hàng chứa tiêu đề
            int currentCol = col_Before + 1; // Bắt đầu quét từ ô bên phải cột "Before"

            while (true)
            {
                string headerValue = ws.Cells[headerRow, currentCol].Text;
                if (string.IsNullOrEmpty(headerValue)) break; // Hết tiêu đề thì dừng

                int cycle = GetFirstNumber(headerValue);
                if (!primeBend)
                {
                    cycle /= 100;
                }

                // Lưu lại ColumnIndex -> Tên cột trong DataTable ("L50", "L100"...)
                cycleColumns.Add(currentCol, $"L{cycle}");
                currentCol++;
            }

            // Lọc danh sách cột cần xử lý
            List<string> listFlexSN = _RESULT[category].Columns.Cast<DataColumn>()
                .Select(col => col.ColumnName)
                .Where(flexSN => !flexSN.Contains("Select") && !flexSN.Contains("Cycle"))
                .ToList();

            int SampleNO = 0, skip = 0;

            // 4. VÒNG LẶP ĐIỀN DỮ LIỆU TỐC ĐỘ CAO (BẰNG TỌA ĐỘ INT)
            foreach (string flexSN in listFlexSN)
            {
                SampleNO++;
                DataTable dt_Detail = GetDetail(category, flexSN);

                if (dt_Detail.Rows.Count > 0)
                {
                    foreach (DataRow row in dt_Detail.Rows)
                    {
                        if (row["Select"]?.ToString().ToUpper() == "TRUE")
                        {
                            skip++;

                            // Chú ý: Tôi giữ nguyên logic skip và skip + 1 của bạn
                            ws.Cells[row_Sample + skip + 1, col_Sample].Value = SampleNO;
                            ws.Cells[row_Net + skip + 1, col_Net].Value = row["Net No"];
                            ws.Cells[row_Flex + skip + 1, col_Flex].Value = row["Flex SN"];
                            ws.Cells[row_Sumi + skip + 1, col_Sumi].Value = row["Test Result"];
                            ws.Cells[row_Cus + skip + 1, col_Cus].Value = row["Net Name"];

                            // Pin1, Pin2, LSL, USL đang cộng "skip" ở code cũ
                            ws.Cells[row_Pin1 + skip, col_Pin1].Value = row["Pin1"];
                            ws.Cells[row_Pin2 + skip, col_Pin2].Value = row["Pin2"];

                            if (double.TryParse(row["LSL"]?.ToString(), out double lsl))
                                ws.Cells[row_LSL + skip, col_LSL].Value = lsl;

                            if (double.TryParse(row["USL"]?.ToString(), out double usl))
                                ws.Cells[row_USL + skip, col_USL].Value = usl;

                            if (double.TryParse(row["BF"]?.ToString(), out double bf))
                                ws.Cells[row_Before + skip + 1, col_Before].Value = bf;

                            // 5. ĐIỀN DỮ LIỆU CYCLE DỰA VÀO BẢN ĐỒ cycleColumns
                            foreach (var kvp in cycleColumns)
                            {
                                int targetCol = kvp.Key;
                                string columnNameInDataTable = kvp.Value;

                                // TryParse sẽ xử lý an toàn giá trị rỗng/null mà không văng Exception
                                if (row.Table.Columns.Contains(columnNameInDataTable) &&
                                    double.TryParse(row[columnNameInDataTable]?.ToString(), out double cycleVal))
                                {
                                    ws.Cells[row_Before + skip + 1, targetCol].Value = cycleVal;
                                }
                            }
                        }
                    }
                }
            }
        }

        static int GetFirstNumber(string input)
        {
            if (string.IsNullOrEmpty(input))
                return 0;

            // Biểu thức chính quy tìm chuỗi chữ số đầu tiên
            Match match = Regex.Match(input, @"\d+");

            if (match.Success && int.TryParse(match.Value, out int result))
            {
                return result;
            }

            return 0; // Trả về null nếu không có số nào trong chuỗi
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

        public string GetError(string cycle, string flexSN, string netNo, string category)
        {
            string key = $"{flexSN}-{cycle}-{switchCategory(category)}-{netNo}";

            if (_ERROR_LIST.TryGetValue(key, out string value))
            {
                return value;
            }

            return "";
        }
    }
}