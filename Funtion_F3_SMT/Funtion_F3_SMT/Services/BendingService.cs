using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using NationalInstruments.Restricted;
using OfficeOpenXml;
using OK2SHIP_SMT.Repositories;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OK2SHIP_SMT.Services
{
    public class BendingService
    {
        private DBContext _dbContext = new DBContext();

        /// <summary>
        /// Danh sách kết quả
        /// {Key: category, value: DataTable (cycle - flexSN)}
        /// </summary>
        private Dictionary<string, DataTable> _RESULT = new Dictionary<string, DataTable>();

        /// <summary>
        /// Spec theo từng 
        /// {Key: category, Value: Spec}
        /// </summary>
        private Dictionary<string, DataTable> __SPEC = new Dictionary<string, DataTable>();

        /// <summary>
        /// Dữ liệu của từng vị trí theo từng hang mục
        /// Key: {flexSN-cycle-category, value: }
        /// </summary>
        private Dictionary<string, string> _DATA = new Dictionary<string, string>();

        /// <summary>
        /// lấy itemcode, lotno, maker từ địa chỉ folder đưa vào
        /// </summary>
        /// <param name="location">địa chỉ folder logfile</param>
        /// <param name="itemCode"></param>
        /// <param name="lotNo"></param>
        /// <param name="maker"></param>
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

        /// <summary>
        /// lấy danh sách kết quả theo tên hạng mục và thêm cột select để người dùng có thể xóa chu kỳ
        /// </summary>
        public DataTable getValue(string category)
        {
            if (_RESULT.TryGetValue(category, out DataTable dt))
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
                string folderName = Path.GetFileName(folderPath);

                if (folderName.Equals("BF", StringComparison.OrdinalIgnoreCase))
                {
                    return -1;
                }

                if (folderName.StartsWith("L", StringComparison.OrdinalIgnoreCase))
                {
                    if (int.TryParse(folderName.Substring(1), out int number))
                    {
                        return number;
                    }
                }

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

        public string switchCategory(string category)
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

            try
            {
                if (__SPEC[category].Rows.Count > 0)
                {
                    if (!dt.Columns.Contains("Net Name"))
                    {
                        DataColumn newColumn = dt.Columns.Add("Net Name", typeof(string));
                        newColumn.SetOrdinal(3);
                        for (int i = 0; i < Math.Min(__SPEC[category].Rows.Count, dt.Rows.Count); i++)
                        {
                            if (__SPEC[category].Columns.Contains("Net Name"))
                            {
                                dt.Rows[i]["Net Name"] = __SPEC[category].Rows[i]["Net Name"];
                            }
                        }
                    }

                    if (!dt.Columns.Contains("Pin1"))
                    {
                        DataColumn newColumn = dt.Columns.Add("Pin1", typeof(string));
                        newColumn.SetOrdinal(4);
                        for (int i = 0; i < Math.Min(__SPEC[category].Rows.Count, dt.Rows.Count); i++)
                        {
                            if (__SPEC[category].Columns.Contains("Pin1"))
                            {
                                dt.Rows[i]["Pin1"] = __SPEC[category].Rows[i]["Pin1"];
                            }
                        }
                    }

                    if (!dt.Columns.Contains("Pin2"))
                    {
                        DataColumn newColumn = dt.Columns.Add("Pin2", typeof(string));
                        newColumn.SetOrdinal(5);
                        for (int i = 0; i < Math.Min(__SPEC[category].Rows.Count, dt.Rows.Count); i++)
                        {
                            if (__SPEC[category].Columns.Contains("Pin2"))
                            {
                                dt.Rows[i]["Pin2"] = __SPEC[category].Rows[i]["Pin2"];
                            }
                        }
                    }

                    if (!dt.Columns.Contains("Select"))
                    {
                        DataColumn newColumn = dt.Columns.Add("Select", typeof(string));
                        newColumn.SetOrdinal(5);
                        for (int i = 0; i < Math.Min(__SPEC[category].Rows.Count, dt.Rows.Count); i++)
                        {
                            if (__SPEC[category].Columns.Contains("Select"))
                            {
                                dt.Rows[i]["Select"] = __SPEC[category].Rows[i]["Select"].ToString().ToUpper() == "TRUE"
                                    ? "YES"
                                    : "NO";
                            }
                        }
                    }
                }
            }
            catch
            {
            }

            return dt;
        }

        private string _NAMETABLE = "BENDING_TEST_DATA";

        public List<string> getListFlexSN(string itemName)
        {
            List<string> result = new List<string>();
            return result;
        }

        public string GetSpecResult()
        {
            List<string> s = new List<string>();
            foreach (string category in _RESULT.Keys)
            {
                //Debugger.Break();
                DataTable dt = _RESULT[category];
                int ngColumnCount = dt.Columns.Cast<DataColumn>()
                    .Count(col => dt.AsEnumerable()
                        .Any(row => row[col] != DBNull.Value && row[col].ToString().Trim().ToUpper() == "NG"));
                s.Add($"{category}\u200B{ngColumnCount}/{_RESULT[category].Columns.Count - 1}");
            }

            return string.Join("\u200F", s);
        }

        public void Save(string itemCode, string lotNo, string maker, int skip)
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
                    List<string> dataZ = dt.Rows[0]["Data"].ToString().Split('\u2060').ToList();
                    foreach (string s in dataZ)
                    {
                        string[] split = s.Split('\u200F');
                        if (!_DATA.TryGetValue(split[0], out string _))
                        {
                            _DATA.Add(split[0], split[1]);
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
            row["Result"] = GetSpecResult();
            if (skip != 1)
            {
                dt.Rows.Add(row);
            }

            _dbContext.BuckDataTable(dt, _NAMETABLE, new[] { "ItemCode", "LotNo", "Maker" }, null, "ID");
        }

        private DataTable SolveResult(string category, bool isSpec = false)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Cycle", typeof(string));

            Dictionary<string, List<string>> _dic = new Dictionary<string, List<string>>();
            foreach (string key in _DATA.Keys)
            {
                string[] keySplit = key.Split('-');
                if (keySplit[2] == category)
                {
                    if (_dic.TryGetValue(keySplit[1], out List<string> list))
                    {
                        list.Add(key);
                    }
                    else
                    {
                        _dic.Add(keySplit[1], new List<string> { key });
                    }
                }
            }

            if (_dic.Count == 0) return dt;

            foreach (string s in _dic[_dic.Keys.First()])
            {
                dt.Columns.Add(s.Split('-')[0]);
            }

            List<string> keys = _dic.Keys
                .ToList()
                .OrderBy(k => k == "BF" ? 0 : 1)
                .ThenBy(k =>
                {
                    if (k.StartsWith("L") && int.TryParse(k.Substring(1), out int num))
                        return num;
                    return int.MaxValue;
                })
                .ToList();


            string[] specSelectArray = null;

            if (isSpec && __SPEC.TryGetValue(switchCategory(category), out DataTable specTable))
            {
                int rowCount = specTable.Rows.Count;
                specSelectArray = new string[rowCount];

                for (int r = 0; r < rowCount; r++)
                {
                    specSelectArray[r] = specTable.Rows[r]["Select"]?.ToString()?.ToUpper();
                }
            }

            foreach (string cycle in keys)
            {
                DataRow row = dt.NewRow();
                row["Cycle"] = cycle;

                foreach (DataColumn col in dt.Columns)
                {
                    string flexSN = col.ColumnName;
                    if (_DATA.TryGetValue($"{flexSN}-{cycle}-{category}", out string data))
                    {
                        string[] listData = data.Split('\u200D');
                        string[] listBF = null;

                        if (dt.Rows.Count > 0 && cycle != "BF")
                        {
                            if (_DATA.TryGetValue($"{flexSN}-BF-{category}", out string bf))
                            {
                                listBF = bf.Split('\u200D');
                            }
                        }

                        string res = "N/A";
                        for (int i = 0; i < listData.Length; i++)
                        {
                            string s = listData[i];
                            if (res == "N/A") res = "OK";

                            string result_err = "OK";
                            string resZ = "";


                            if (isSpec && specSelectArray != null && i < specSelectArray.Length)
                            {
                                resZ = specSelectArray[i];
                                if (string.IsNullOrEmpty(resZ))
                                {
                                    Debugger.Break();
                                }

                                resZ = resZ == "TRUE" ? "YES" : resZ;
                            }


                            if (resZ == "YES")
                            {
                                string[] splitS = s.Split('\u200B');
                                if (double.TryParse(splitS[2], out double low) &&
                                    double.TryParse(splitS[3], out double high) &&
                                    double.TryParse(splitS[0], out double measure))
                                {
                                    if (measure < low || measure > high)
                                    {
                                        res = "NG";
                                        result_err = low > measure ? "Over USL" : "Over LSL";
                                    }

                                    if (listBF != null && i < listBF.Length)
                                    {
                                        string[] splitBF = listBF[i].Split('\u200B');
                                        if (double.TryParse(splitBF[0], out double valueBF))
                                        {
                                            double z = (valueBF - measure) / valueBF;
                                            bool check = true;


                                            string testResult = splitS[1];
                                            if (testResult != null && testResult.Contains("SUS"))
                                            {
                                                check = false;
                                            }


                                            if (check)
                                            {
                                                // viết biến thiên custom ở đây
                                                var table = __SPEC[switchCategory(category)];
                                                string specCustomer = table.Columns.Contains("Custom Spec") &&
                                                                      !table.Rows[i].IsNull("Custom Spec")
                                                    ? table.Rows[i]["Custom Spec"].ToString()
                                                    : string.Empty;

                                                // mặc định spec là 10% - theo sumitomo
                                                double checkAgain = 0.1;

                                                // mặc định spec là 500% đối với các mục - theo sumitomo

                                                if (category.Contains("TC") || category.Contains("TS") ||
                                                    category.Contains("HS"))
                                                {
                                                    checkAgain = 5;
                                                }

                                                if (!string.IsNullOrEmpty(specCustomer))
                                                {
                                                    string[] split = specCustomer.Split(';');
                                                    foreach (string itemCycle in split)
                                                    {
                                                        string[] sList = itemCycle.Split(':');
                                                        if (sList[0].Contains(cycle))
                                                        {
                                                            if (double.TryParse(sList[1], out double value))
                                                            {
                                                                checkAgain = value;
                                                            }
                                                        }
                                                    }
                                                }


                                                if (z > checkAgain || z < -checkAgain)
                                                {
                                                    res = "NG";
                                                    result_err = "Vanability R NG";
                                                }
                                                else if (z > 0.095 || z < -0.095)
                                                {
                                                    result_err = "Vanability T NG";
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    res = "N/A";
                                    result_err = "N/A";
                                }
                            }


                            _ERROR_LIST[$"{flexSN}-{cycle}-{category}-{i}"] = result_err;
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
            __SPEC.Clear();
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

            List<string> data = dt.Rows[0]["Data"].ToString().Split('\u2060').ToList();
            Dictionary<string, string> dic = new Dictionary<string, string>();

            foreach (string s in data)
            {
                string[] split = s.Split('\u200F');
                _DATA.Add(split[0], split[1]);
                string category = split[0].Split('-')[2];
                if (!dic.TryGetValue(category, out string value))
                {
                    dic.Add(category, "1");
                }
            }


            foreach (string s in dic.Keys)
            {
                _RESULT.Add(switchCategory(s), SolveResult(s, true));
            }
        }

        public void MergeExcelFiles(string[] sourceFiles, string destinationFile)
        {
            // 1. Khởi tạo file đích
            using (SpreadsheetDocument destDoc =
                   SpreadsheetDocument.Create(destinationFile, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart destWorkbookPart = destDoc.AddWorkbookPart();
                destWorkbookPart.Workbook = new Workbook();
                Sheets destSheets = destDoc.WorkbookPart.Workbook.AppendChild(new Sheets());

                // Tạo bảng chứa chuỗi (SharedStringTable) cho file đích
                SharedStringTablePart destStringTablePart = destWorkbookPart.AddNewPart<SharedStringTablePart>();
                destStringTablePart.SharedStringTable = new SharedStringTable();

                uint sheetId = 1;
                bool isFirstFile = true;

                // Cấu trúc dữ liệu để cache SharedStrings, giúp tăng tốc đáng kể
                var globalStringCache = new Dictionary<string, int>();

                // 2. Lặp qua danh sách file
                foreach (var file in sourceFiles)
                {
                    Console.WriteLine($"Dang xu ly: {Path.GetFileName(file)}");

                    using (SpreadsheetDocument srcDoc = SpreadsheetDocument.Open(file, false))
                    {
                        WorkbookPart srcWorkbookPart = srcDoc.WorkbookPart;
                        SharedStringTablePart srcStringTablePart =
                            srcWorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                        // 3. Copy bộ Style (định dạng ô, màu sắc) từ file đầu tiên
                        if (isFirstFile)
                        {
                            WorkbookStylesPart srcStylesPart =
                                srcWorkbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
                            if (srcStylesPart != null)
                            {
                                WorkbookStylesPart destStylesPart = destWorkbookPart.AddNewPart<WorkbookStylesPart>();
                                destStylesPart.FeedData(srcStylesPart.GetStream());
                            }

                            isFirstFile = false;
                        }

                        // 4. Lặp qua từng Sheet trong file nguồn
                        foreach (Sheet srcSheet in srcWorkbookPart.Workbook.Sheets.Elements<Sheet>())
                        {
                            WorksheetPart srcWorksheetPart = (WorksheetPart)srcWorkbookPart.GetPartById(srcSheet.Id);

                            // Clone toàn bộ Worksheet (kéo theo cả Hình ảnh và Drawings)
                            WorksheetPart clonedSheetPart = destWorkbookPart.AddPart(srcWorksheetPart);
                            string relationId = destWorkbookPart.GetIdOfPart(clonedSheetPart);

                            // 5. Remap lại dữ liệu chữ (Shared Strings)
                            if (srcStringTablePart != null)
                            {
                                // Ánh xạ Old Index -> New Index chỉ cho Sheet hiện tại
                                var localStringDict = new Dictionary<int, int>();

                                // Tìm tất cả các ô chứa dữ liệu chuỗi (SharedString)
                                var cells = clonedSheetPart.Worksheet.Descendants<Cell>()
                                    .Where(c => c.DataType != null && c.DataType.Value == CellValues.SharedString);

                                foreach (Cell cell in cells)
                                {
                                    int oldIndex = int.Parse(cell.CellValue.Text);

                                    if (!localStringDict.ContainsKey(oldIndex))
                                    {
                                        // Lấy text thực tế từ file nguồn
                                        SharedStringItem oldItem = srcStringTablePart.SharedStringTable
                                            .Elements<SharedStringItem>().ElementAt(oldIndex);
                                        string realText = oldItem.InnerText;

                                        // Kiểm tra xem text này đã tồn tại trong file tổng chưa (để tránh phình dung lượng)
                                        if (!globalStringCache.ContainsKey(realText))
                                        {
                                            int newIndex = globalStringCache.Count;
                                            destStringTablePart.SharedStringTable.AppendChild(
                                                new SharedStringItem(new Text(realText)));
                                            globalStringCache[realText] = newIndex;
                                        }

                                        localStringDict[oldIndex] = globalStringCache[realText];
                                    }

                                    // Cập nhật lại Cell với Index mới của file đích
                                    cell.CellValue.Text = localStringDict[oldIndex].ToString();
                                }

                                // Lưu lại thay đổi của Sheet vừa remap
                                clonedSheetPart.Worksheet.Save();
                            }

                            // 6. Đăng ký Sheet vào file đích
                            string safeSheetName = TruncateSheetName($"{srcSheet.Name}_{sheetId}");
                            Sheet newSheet = new Sheet() { Id = relationId, SheetId = sheetId, Name = safeSheetName };
                            destSheets.Append(newSheet);
                            sheetId++;
                        }
                    }
                }

                // Lưu toàn bộ file tổng hợp
                destStringTablePart.SharedStringTable.Save();
                destWorkbookPart.Workbook.Save();
            }
        }

        // Hàm phụ trợ: Excel giới hạn tên Sheet tối đa 31 ký tự
        private static string TruncateSheetName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Sheet";
            return name.Length > 31 ? name.Substring(0, 31) : name;
        }


        public void MergeFile(string[] listFileSave, string nameDIC)
        {
            try
            {
                string dirOutPut = new ExportProcess().getLocationSave();

                dirOutPut = Path.Combine(dirOutPut, $"ORT\\{nameDIC}.xlsx");
                string folderPath = Path.GetDirectoryName(dirOutPut);

                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                if (false)
                {
                    // mở khóa sức mạnh tiềm ẩn
                    MergeExcelFiles(listFileSave, dirOutPut);
                    return;
                }

                // 1. Lấy đường dẫn động tới thư mục chứa phần mềm
                string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

                // Tạo đường dẫn chuẩn đến file .exe
                string exePath = Path.Combine(baseDirectory, "MergeService", "MergeHeavyExcelService.exe");

                // Kiểm tra xem file .exe có tồn tại không
                if (!File.Exists(exePath))
                {
                    throw new Exception("Không tìm thấy file dịch vụ MergeHeavyExcelService.exe tại: " + exePath);
                }

                // 2. Tạo chuỗi arguments (tham số)
                StringBuilder arguments = new StringBuilder();

                // Nối danh sách các file input (Mỗi đường dẫn bọc trong dấu ngoặc kép "")
                foreach (string file in listFileSave)
                {
                    arguments.Append($"\"{file}\" ");
                }

                // Nối file output vào cuối cùng (Bọc trong ngoặc kép "")
                arguments.Append($"\"{dirOutPut}\"");

                // 3. Cấu hình tiến trình (Process) chạy file exe
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = arguments.ToString(),
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    // THÊM: Đọc luồng để bắt lỗi từ Python
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };

                // 4. Khởi chạy tiến trình
                using (Process process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();

                    // Đọc lại toàn bộ nội dung mà Python đã print() ra
                    string outputLog = process.StandardOutput.ReadToEnd();
                    string errorLog = process.StandardError.ReadToEnd();

                    process.WaitForExit();

                    // 5. Kiểm tra mã trả về
                    if (process.ExitCode != 0)
                    {
                        // Bắn ra Exception kèm theo log của file Python để dễ debug
                        throw new Exception(
                            $"Dịch vụ ghép file thất bại (Mã lỗi: {process.ExitCode}).\n\nLog hệ thống:\n{outputLog}\n{errorLog}");
                    }
                }

                // 6. MỞ FILE SAU KHI GHÉP THÀNH CÔNG
                Process.Start(new ProcessStartInfo(dirOutPut)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                // Ném lỗi lên hàm gọi nó (để UI có thể catch và show MessageBox nếu cần)
                throw new Exception("Lỗi trong quá trình MergeFile: " + ex.Message);
            }
        }

        public string Export(string itemCode, string lotNo, string maker, string category = "Flex Bending",
            bool lockZ = true, bool ers = true)
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
                    namesheet = "Thermal Shock";
                    ictsheet = "Thermal Shock - ICT data";
                    break;

                default:
                    return "";
            }

            if (string.IsNullOrEmpty(maker))
            {
                DataTable dt_maker = _dbContext.LoadDataTable(_NAMETABLE, new[] { "ItemCode", "LotNo" },
                    new[] { itemCode, lotNo }, new[] { "ItemCode", "LotNo", "Maker" });
                ExportProcess process = new ExportProcess();

                using (ExcelPackage package = process.FindFormatProcess("", itemCode, lotNo))
                {
                    List<string> name = new[] { namesheet, ictsheet }.ToList();

                    using (ExcelWorksheet worksheet = process.FindSheetWithKey(package, namesheet))
                    using (ExcelWorksheet worksheet_ICT = process.FindSheetWithKey(package, ictsheet))
                    using (ExcelWorksheet worksheetXray = process.FindSheetWithKey(package, xraysheet))
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

                        Dictionary<string, string> flexSNDIC = new Dictionary<string, string>();
                        List<string> makers = new List<string>();
                        foreach (DataRow row in dt_maker.Rows)
                        {
                            itemCode = row["ItemCode"].ToString();
                            lotNo = row["LotNo"].ToString();
                            maker = row["Maker"].ToString();

                            Load(itemCode, lotNo, maker);
                            List<string> listFlexSN = _RESULT[category].Columns.Cast<DataColumn>()
                                .Select(col => col.ColumnName)
                                .Where(flexSN => !flexSN.Contains("Select") && !flexSN.Contains("Cycle"))
                                .ToList();
                            flexSNDIC.Add(maker, string.Join("-", listFlexSN));
                            if (primeBend)
                            {
                                DataTable dt = _RESULT[category];
                                ExportDataToWS(worksheet, category, dt, maker, slot);

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

                            ExcelWorksheet currentIctSheet = worksheet_ICT;

                            if (slot > 0)
                            {
                                string newIctSheetName = $"{ictsheet}{slot}";
                                currentIctSheet = package.Workbook.Worksheets.Add(newIctSheetName, worksheet_ICT);
                                name.Add(newIctSheetName);
                            }

                            ExportDataToICT(currentIctSheet, category, primeBend, ers);
                            slot++;
                        }

                        try
                        {
                            if (makers.Count > 0)
                            {
                                XRayPictureService service = new XRayPictureService();
                                service.ExportMultipleMakers(worksheetXray, itemCode, lotNo, makers, flexSNDIC,
                                    category);
                            }
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            Debugger.Break();
#endif
                        }

                        package.Compression = OfficeOpenXml.CompressionLevel.BestSpeed;
                        return process.SaveExcelWorksheet(package, string.Join(":", name),
                            $"{itemCode}-{lotNo}-{category}", "NPI", lockZ);
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
                                List<string> listFlexSN = _RESULT[category].Columns.Cast<DataColumn>()
                                    .Select(col => col.ColumnName)
                                    .Where(flexSN => !flexSN.Contains("Select") && !flexSN.Contains("Cycle"))
                                    .ToList();
                                service.Export(worksheetXray, itemCode, lotNo, maker, category, false, 0, listFlexSN);
                            }
                            catch
                            {
                                //Debugger.Break();
                            }
                        }
                        else
                        {
                            ExportDataToWSNotBend(worksheet, category, 0);
                        }

                        ExportDataToICT(worksheet_ICT, category, primeBend, ers);
                        package.Compression = OfficeOpenXml.CompressionLevel.BestSpeed;
                        return process.SaveExcelWorksheet(package, string.Join(":", name), $"{itemCode}-{lotNo}", "NPI",
                            lockZ);
                    }
                }
            }
        }

        public void checkERS(string itemCode, string lotNo, string maker)
        {
            try
            {
                // kiểm tra có tồn tại hay không
                ERS_NetSpecService service = new ERS_NetSpecService();
                service.LoadData(itemCode, maker);

                int numSpec = service.getNumberOfData();

                DataTable dt = _dbContext.LoadDataTable(_NAMETABLE, new[] { "ItemCode", "LotNo", "Maker" },
                    new[] { itemCode, lotNo, maker });

                if (dt.Rows.Count <= 0)
                {
                    throw new DataException("Data is empty");
                }

                List<string> data = dt.Rows[0]["Data"].ToString().Split('\u2060').ToList();
                Dictionary<string, string> dic = new Dictionary<string, string>();

                foreach (string s in data)
                {
                    string[] split = s.Split('\u200F');
                    int z = split[1].Split('\u200D').Length;
                    if (numSpec != z)
                    {
                        string[] d = split[0].Split('-');

                        throw new Exception($"Number of Ers not match data Flex {d[0]} - cycle {d[1]}");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Check ERS: {ex.Message}");
            }
        }

        private void SetupXray(ExcelWorksheet worksheet, int slot)
        {
            if (worksheet == null)
            {
                throw new Exception("worksheet xray is not valid");
            }

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
            if (worksheet == null)
            {
                throw new Exception("worksheet is not valid");
            }

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
                                if (row["Cycle"] != DBNull.Value && row["Cycle"].ToString() == cycleZ.ToString())
                                {
                                    value = row[dt.Columns[i]].ToString();
                                    break;
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
                    i++;
                }
            }
            catch
            {
                //Debugger.Break();
            }
        }


        private void ExportDataToWS(ExcelWorksheet ws, string category, DataTable dt, string maker, int slot = 0)
        {
            IDictionary<string, string> _dic =
                ExportProcess.FindAddressByText(ws, new[] { "Flex SN", "E-check", "Sample", "Condition" });


            string[] splitSample = _dic["Sample"].Split('-');
            List<string> cache = new List<string>();
            string[] splitCondition = _dic["Condition"].Split('-');
            ws.Cells[ExportProcess.AddColumn(splitCondition[slot], 1)].Value = maker;
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

        private void ExportDataToICT(ExcelWorksheet ws, string category, bool primeBend = false, bool isSus = true)
        {
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(ws, new[]
            {
                "Sample no.", "Net no.", "Flex SN", "Sumitomo net name", "Customer net name", "Pin 1", "Pin 2",
                "LSL", "USL", "Before", "% Resistance change", "Judgement"
            });

            dic["Before"] = dic["Before"].Split('-').Last();

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

            int col_Judge = -1;
            if (dic.ContainsKey("Judgement"))
            {
                col_Judge = ws.Cells[dic["Judgement"].Split('-').Last()].Start.Column;
            }

            Dictionary<string, int> rawCycleCols = new Dictionary<string, int>();

            // CỘNG THÊM "BF" VÀO rawCycleCols ĐỂ BẢNG SUMMARY TÍNH ĐƯỢC CỘT BEFORE
            rawCycleCols["BF"] = col_Before;
            int currentCol = col_Before + 1;

            while (true)
            {
                string headerValue = ws.Cells[row_Before, currentCol].Text;
                if (string.IsNullOrEmpty(headerValue) || headerValue.Contains("%") ||
                    headerValue.ToUpper().Contains("RESISTANCE"))
                    break;

                int cycle = GetFirstNumber(headerValue);
                if (cycle > 0)
                {
                    if (!primeBend) cycle /= 100;
                    rawCycleCols[$"L{cycle}"] = currentCol;
                }
                else break;

                currentCol++;
            }

            Dictionary<string, int> percentCycleCols = new Dictionary<string, int>();
            if (dic.ContainsKey("% Resistance change"))
            {
                string percentAddr = dic["% Resistance change"].Split('-').Last();
                int percentStartCol = ws.Cells[percentAddr].Start.Column;

                currentCol = percentStartCol;
                while (true)
                {
                    string headerValue = ws.Cells[row_Before, currentCol].Text;
                    if (string.IsNullOrEmpty(headerValue) || headerValue.ToUpper().Contains("JUDGEMENT"))
                        break;

                    int cycle = GetFirstNumber(headerValue);
                    if (cycle > 0)
                    {
                        if (!primeBend) cycle /= 100;
                        percentCycleCols[$"L{cycle}"] = currentCol;
                    }
                    else break;

                    currentCol++;
                }
            }

            List<string> listFlexSN = _RESULT[category].Columns.Cast<DataColumn>()
                .Select(col => col.ColumnName)
                .Where(flexSN => !flexSN.Contains("Select") && !flexSN.Contains("Cycle"))
                .ToList();

            int SampleNO = 0, skip = 0;

            Dictionary<string, List<double>> allNetValues = new Dictionary<string, List<double>>();
            Dictionary<string, List<double>> allNetPercents = new Dictionary<string, List<double>>();
            Dictionary<string, List<double>> susNetValues = new Dictionary<string, List<double>>();
            Dictionary<string, List<double>> susNetPercents = new Dictionary<string, List<double>>();
            int lastRow = 0;
            foreach (string flexSN in listFlexSN)
            {
                SampleNO++;
                int netNo = 1;
                DataTable dt_Detail = GetDetail(category, flexSN);

                if (dt_Detail.Rows.Count > 0)
                {
                    foreach (DataRow row in dt_Detail.Rows)
                    {
                        string selectVal;
                        try
                        {
                            selectVal = row["Select"]?.ToString().ToUpper();
                        }
                        catch
                        {
                            selectVal = "";
                        }

                        // kiểm tra xem có ers không
                        if (string.IsNullOrEmpty(selectVal))
                        {
                            if (isSus)
                            {
                                selectVal = "YES";
                            }
                            else
                            {
                                throw new Exception("Ers is not setup!");
                            }
                        }

                        if (selectVal == "YES" || selectVal == "TRUE" || isSus)
                        {
                            skip++;
                            int targetRow = row_Sample + skip + 1;
                            lastRow = Math.Max(targetRow, lastRow);
                            ws.Cells[targetRow, col_Sample].Value = SampleNO;
                            ws.Cells[targetRow, col_Net].Value = netNo++;
                            ws.Cells[targetRow, col_Flex].Value = row["Flex SN"];
                            ws.Cells[targetRow, col_Sumi].Value = row["Test Result"];
                            try
                            {
                                ws.Cells[targetRow, col_Cus].Value = row["Net Name"];

                                ws.Cells[row_Pin1 + skip, col_Pin1].Value = row["Pin1"];
                                ws.Cells[row_Pin2 + skip, col_Pin2].Value = row["Pin2"];
                                ws.Cells[row_Pin2 + skip, col_Pin2 + 1].Value = "yes";
                            }
                            catch
                            {
                            }

                            if (double.TryParse(row["LSL"]?.ToString(), out double lsl))
                                ws.Cells[row_LSL + skip, col_LSL].Value = lsl;
                            if (double.TryParse(row["USL"]?.ToString(), out double usl))
                                ws.Cells[row_USL + skip, col_USL].Value = usl;

                            double bf = 0;
                            bool hasBf = false;
                            if (double.TryParse(row["BF"]?.ToString(), out double parsedBf))
                            {
                                bf = parsedBf;
                                hasBf = true;
                            }

                            bool isOver10Percent = false;
                            bool hasCalc = false;

                            string netNameVal = "", sumiNameVal = "";
                            try
                            {
                                netNameVal = row["Net Name"]?.ToString().ToUpper();
                                sumiNameVal = row["Test Result"]?.ToString().ToUpper() ?? "";
                            }
                            catch
                            {
                            }

                            // Phân loại: Netname có chứa chữ SUS
                            bool isRowSus = netNameVal.Contains("SUS") || sumiNameVal.Contains("SUS");
                            foreach (var kvp in rawCycleCols)
                            {
                                string colName = kvp.Key;
                                int rawColIndex = kvp.Value;

                                if (row.Table.Columns.Contains(colName) &&
                                    double.TryParse(row[colName]?.ToString(), out double cycleVal))
                                {
                                    ws.Cells[targetRow, rawColIndex].Value = cycleVal;

                                    if (isRowSus)
                                    {
                                        if (!susNetValues.ContainsKey(colName))
                                            susNetValues[colName] = new List<double>();
                                        susNetValues[colName].Add(cycleVal);
                                    }
                                    else
                                    {
                                        if (!allNetValues.ContainsKey(colName))
                                            allNetValues[colName] = new List<double>();
                                        allNetValues[colName].Add(cycleVal);
                                    }

                                    if (hasBf && percentCycleCols.TryGetValue(colName, out int percentColIndex))
                                    {
                                        var percentCell = ws.Cells[targetRow, percentColIndex];
                                        percentCell.Formula = "";

                                        if (bf != 0)
                                        {
                                            double percentChange = (cycleVal - bf) / bf;
                                            percentCell.Value = percentChange;
                                            percentCell.Style.Numberformat.Format = "0.00%";
                                            hasCalc = true;

                                            if (isRowSus)
                                            {
                                                if (!susNetPercents.ContainsKey(colName))
                                                    susNetPercents[colName] = new List<double>();
                                                susNetPercents[colName].Add(percentChange);
                                            }
                                            else
                                            {
                                                if (!allNetPercents.ContainsKey(colName))
                                                    allNetPercents[colName] = new List<double>();
                                                allNetPercents[colName].Add(percentChange);
                                            }

                                            //Debugger.Break();
                                            //sumitomo
                                            double specCustom = 0.1;
                                            string c = switchCategory(category);
                                            if (c == "TS" || c == "HS" || c == "TC")
                                            {
                                                specCustom = 5;
                                            }

                                            string iitem = row["Net No"].ToString();
                                            int i = int.Parse(iitem);
                                            // // viết biến thiên custom ở đây
                                            var table = __SPEC[category];

                                            string specCustomer = "";
                                            try
                                            {
                                                if (table.Rows.Count > i)
                                                {
                                                    specCustomer = table.Columns.Contains("Custom Spec") &&
                                                                   !table.Rows[i].IsNull("Custom Spec")
                                                        ? table.Rows[i]["Custom Spec"].ToString()
                                                        : string.Empty;
                                                }
                                            }
                                            catch
                                            {
                                            }


                                            if (!string.IsNullOrEmpty(specCustomer))
                                            {
                                                string[] split = specCustomer.Split(';');
                                                foreach (string itemCycle in split)
                                                {
                                                    string[] sList = itemCycle.Split(':');
                                                    if (sList[0].Contains(kvp.Key))
                                                    {
                                                        if (double.TryParse(sList[1], out double value))
                                                        {
                                                            specCustom = value;
                                                        }
                                                    }
                                                }
                                            }

                                            if (percentChange > specCustom || percentChange < -specCustom)
                                            {
                                                isOver10Percent = true;
                                            }
                                        }
                                        else
                                        {
                                            percentCell.Value = "";
                                        }
                                    }
                                }
                            }

                            if (col_Judge != -1)
                            {
                                var judgeCell = ws.Cells[targetRow, col_Judge];
                                judgeCell.Formula = "";

                                if (hasCalc)
                                {
                                    judgeCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

                                    if (!isOver10Percent)
                                    {
                                        judgeCell.Value = "Pass-resistance change is within ±10%.";


                                        judgeCell.Style.Fill.BackgroundColor.SetColor(
                                            System.Drawing.Color.FromArgb(198, 239, 206));
                                        judgeCell.Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(0, 97, 0));
                                    }
                                    else
                                    {
                                        bool isBending = category.ToUpper().Contains("BEND");
                                        if (isBending && !isRowSus)
                                        {
                                            judgeCell.Value = "Resistance change over ±10%";
                                            judgeCell.Style.Fill.BackgroundColor.SetColor(
                                                System.Drawing.Color.FromArgb(255, 235, 156));
                                            judgeCell.Style.Font.Color.SetColor(
                                                System.Drawing.Color.FromArgb(156, 87, 0));
                                        }
                                        else
                                        {
                                            if (isRowSus)
                                            {
                                                judgeCell.Value =
                                                    "Pass following Sumitomo spec-resistance change is outside ±10%.";
                                            }
                                            else
                                            {
                                                if (category == "Heat soak" || category == "Thermal shock" ||
                                                    category == "Thermal cycling")
                                                {
                                                    judgeCell.Value =
                                                        "Pass following Sumitomo spec-resistance change is outside ±10%.";
                                                }
                                                else
                                                {
                                                    judgeCell.Value =
                                                        "Fail-resistance change is within ±10%";
                                                }
                                            }

                                            judgeCell.Style.Fill.BackgroundColor.SetColor(
                                                System.Drawing.Color.FromArgb(198, 239, 206));
                                            judgeCell.Style.Font.Color.SetColor(
                                                System.Drawing.Color.FromArgb(0, 97, 0));
                                        }
                                    }
                                }
                                else
                                {
                                    judgeCell.Value = "";
                                    judgeCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.None;
                                }
                            }
                        }
                    }
                }
            }

            try
            {
                lastRow = lastRow + 1;
                int lastRowZ = ws.Dimension.End.Row;
                ws.DeleteRow(lastRow, lastRowZ - lastRow + 1);
            }
            catch
            {
            }

            // ==============================================================================
            // 5. ĐIỀN BẢNG SUMMARY VÀO TEMPLATE ĐÃ CÓ SẴN (KHÔNG TỰ ĐỘNG SINH HÀNG)
            // ==============================================================================
            try
            {
                int summaryRowStart = -1;
                int susRowStart = -1;

                // Quét 100 dòng đầu tiên để lấy tọa độ dòng "Max" của 2 bảng All nets & SUS net
                int maxRowCheck = Math.Min(ws.Dimension?.End.Row ?? 100, 80);
                for (int r = 1; r <= maxRowCheck; r++)
                {
                    for (int c = 1; c <= 15; c++)
                    {
                        string cellText = ws.Cells[r, c].Text?.ToUpper() ?? "";

                        // Tìm bảng All nets (Có chữ ALL NETS và có chữ EXCLUDE để đảm bảo lấy đúng cụm từ "exclude SUS net")
                        if (summaryRowStart == -1 && cellText.Contains("ALL NETS") && cellText.Contains("EXCLUDE"))
                        {
                            summaryRowStart = r;
                        }

                        // Tìm bảng SUS net (Có chữ SUS NET nhưng TUYỆT ĐỐI KHÔNG ĐƯỢC có chữ EXCLUDE)
                        if (susRowStart == -1 && cellText.Contains("SUS NET") && !cellText.Contains("EXCLUDE"))
                        {
                            susRowStart = r;
                        }
                    }
                }

                void FillSummaryBlock(int startRow, Dictionary<string, List<double>> valDict,
                    Dictionary<string, List<double>> perDict)
                {
                    if (startRow == -1) return;

                    // Xóa rác và công thức cũ (nếu có)
                    for (int r = startRow; r <= startRow + 2; r++)
                    {
                        foreach (var colIdx in rawCycleCols.Values) ws.Cells[r, colIdx].Formula = "";
                        foreach (var colIdx in percentCycleCols.Values) ws.Cells[r, colIdx].Formula = "";
                    }

                    // Điền data giá trị (Resistance value summary)
                    foreach (var kvp in rawCycleCols)
                    {
                        string cycleKey = kvp.Key;
                        int colIdx = kvp.Value;

                        if (valDict.ContainsKey(cycleKey) && valDict[cycleKey].Count > 0)
                        {
                            var list = valDict[cycleKey];
                            ws.Cells[startRow, colIdx].Value = Math.Round(list.Max(), 3);
                            ws.Cells[startRow + 1, colIdx].Value = Math.Round(list.Min(), 3);
                            ws.Cells[startRow + 2, colIdx].Value = Math.Round(list.Average(), 3);
                        }
                        else
                        {
                            ws.Cells[startRow, colIdx].Value = "";
                            ws.Cells[startRow + 1, colIdx].Value = "";
                            ws.Cells[startRow + 2, colIdx].Value = "";
                        }
                    }

                    // Điền data % (% Resistance change summary)
                    foreach (var kvp in percentCycleCols)
                    {
                        string cycleKey = kvp.Key;
                        int colIdx = kvp.Value;

                        if (perDict.ContainsKey(cycleKey) && perDict[cycleKey].Count > 0)
                        {
                            var list = perDict[cycleKey];

                            var maxCell = ws.Cells[startRow, colIdx];
                            maxCell.Value = list.OrderByDescending(x => Math.Abs(x)).First();
                            maxCell.Style.Numberformat.Format = "0.00%";

                            var minCell = ws.Cells[startRow + 1, colIdx];
                            minCell.Value = list.OrderByDescending(x => Math.Abs(x)).Last();
                            minCell.Style.Numberformat.Format = "0.00%";

                            var averCell = ws.Cells[startRow + 2, colIdx];
                            averCell.Value = list.Average();
                            averCell.Style.Numberformat.Format = "0.00%";
                        }
                        else
                        {
                            ws.Cells[startRow, colIdx].Value = "";
                            ws.Cells[startRow + 1, colIdx].Value = "";
                            ws.Cells[startRow + 2, colIdx].Value = "";
                        }
                    }
                }

                // 1. Điền bảng All nets Summary (Bỏ qua SUS net)
                if (summaryRowStart != -1)
                {
                    FillSummaryBlock(summaryRowStart, allNetValues, allNetPercents);
                }

                // 2. Điền bảng SUS net Summary (Chỉ gồm các net có chứa SUS)
                if (susRowStart != -1)
                {
                    FillSummaryBlock(susRowStart, susNetValues, susNetPercents);
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                Debugger.Break();
#endif
            }
        }

        static int GetFirstNumber(string input)
        {
            if (string.IsNullOrEmpty(input))
                return 0;

            Match match = Regex.Match(input, @"\d+");

            if (match.Success && int.TryParse(match.Value, out int result))
            {
                return result;
            }

            return 0;
        }

        public DataTable GetDetailByCycle(string category, string cycle)
        {
            cycle = cycle.Trim();
            DataTable dt = new DataTable();
            dt.Columns.Add("NetNo", typeof(string));
            Dictionary<string, string> FlexSNList = new Dictionary<string, string>();

            foreach (string key in _DATA.Keys)
            {
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

        public string GetDelta(string netNoValue, string flexSnValue,
            string cycle, string category, double valuez)
        {
            string key = $"{flexSnValue}-BF-{switchCategory(category)}";

            if (_DATA.TryGetValue(key, out string value) && int.TryParse(netNoValue, out int i))
            {
                if (i > 0)
                {
                    List<string> z = value.Split('\u200D').ToList();
                    string valueBF = z[i - 1].Split('\u200B')[0];
                    if (double.TryParse(valueBF, out double valueN))
                    {
                        double Hz = Math.Abs((valuez - valueN) / valueN * 100);
                        string delta = $"Before Value = {valueBF}\n △R = {Hz:F2}%";
                        return delta;
                    }
                }
            }

            return "";
        }
    }
}