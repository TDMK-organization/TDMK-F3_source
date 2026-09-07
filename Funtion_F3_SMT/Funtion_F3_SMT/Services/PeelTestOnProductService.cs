using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.Services
{
    class PeelTestOnProductService
    {
        private DBContext _dbContext = new DBContext();
        public string _ITEMCODE { get; set; }
        public string _LOTNO { get; set; }
        public string _TYPE { get; set; }
        private DBContext DBContext = new DBContext();
        public Dictionary<string, DataTable> _SPEC = new Dictionary<string, DataTable>();
        public Dictionary<string, DataTable> _DIC = new Dictionary<string, DataTable>();
        public Dictionary<string, DataTable> _BEFOREIMAGE = new Dictionary<string, DataTable>();
        private DataTable SetUpDT()
        {
            DataTable dataTable = new DataTable();
            dataTable = new DataTable();
            dataTable.Columns.Add("Id", typeof(int));
            dataTable.Columns.Add("ItemCode");
            dataTable.Columns.Add("LotNo");
            dataTable.Columns.Add("Picture", typeof(Image));
            return dataTable;

        }

        public void CheckSpec(string key)
        {
            DataTable spec = _SPEC[$"{_ITEMCODE} - {_LOTNO}"];
            DataRow r = null;
            foreach (DataRow row in spec.Rows)
            {
                if (key.Contains(row["Type"].ToString().Trim()) && key.Contains(row["Tape"].ToString().Trim()))
                {
                    r = row;
                    break;
                }
            }
            try
            {

                double maxPeakN = double.Parse(r["Peak(N)"].ToString().Split('~')[1]);
                double minPeakN = double.Parse(r["Peak(N)"].ToString().Split('~')[0]);
                double maxAveN = double.Parse(r["Average(N)"].ToString().Split('~')[1]);
                double minAveN = double.Parse(r["Average(N)"].ToString().Split('~')[0]);

                foreach (DataRow row in _DIC[key].Rows)
                {
                    bool prime = true;
                    if (double.TryParse(row["Peak(N)"].ToString().Trim(), out double num) && (num <= maxPeakN && minPeakN <= num))
                    {
                    }
                    else
                    {
                        prime = false;
                    }

                    if (double.TryParse(row["Average(N)"].ToString().Trim(), out num) && (num <= maxAveN && minAveN <= num))
                    {
                    }
                    else
                    {
                        prime = false;
                    }

                    if (prime)
                    {
                        row["JudgementForce"] = "OK";
                    }
                    else
                    {
                        row["JudgementForce"] = "NG";
                    }

                }
            }
            catch
            {
                return;
            }
        }
        public void MakerSpec(string itemCode, string lotNo)
        {
            //List<DataTable> listDt = new List<DataTable>();
            //listDt.Add(_dbContext.LoadDataTable("SPEC_COMMENT_3", new[] { "ItemCode", "Sheet", "Remark" }, new[] { itemCode, "LINER_PEEL_TEST_ON_PRODUCT", "NPI" }));
            //listDt.Add(_dbContext.LoadDataTable("SPEC_COMMENT_3", new[] { "ItemCode", "Sheet", "Remark" }, new[] { itemCode, "PSA_PEEL_TEST_ON_PRODUCT", "NPI" }));
            string _key = $"{itemCode} - {lotNo}";
            _SPEC.Add(_key, setUpSpec());
            //Dictionary<string, DataRow> dic = new Dictionary<string, DataRow>();

            //foreach (DataTable dataTable in listDt)
            //{
            //    foreach (DataRow row in dataTable.Rows)
            //    {
            //        string type = row["Sheet"].ToString().Split('_')[0];
            //        string counting = row["Count_Sample"].ToString();
            //        string[] lists = row["Location"].ToString().Split('-');
            //        foreach (string item in lists)
            //        {
            //            //Debugger.Break();
            //            string[] iteZ = item.Split('|');
            //            if (iteZ.Count() >= 2)
            //            {

            //                if (dic.TryGetValue($"{type}-{iteZ[0]}", out DataRow valueRow))
            //                {

            //                }
            //                else
            //                {
            //                    valueRow = _SPEC[_key].NewRow();
            //                    valueRow["ItemCode"] = itemCode;
            //                    valueRow["LotNo"] = lotNo;
            //                    valueRow["Type"] = type;
            //                    valueRow["Tape"] = iteZ[0];
            //                }
            //                try
            //                {

            //                    if (iteZ[1].Contains("Average"))
            //                    {
            //                        if (iteZ[1].Contains("(gf)"))
            //                        {
            //                            valueRow["Average(Gf)"] = iteZ[1].Split(':')[1];
            //                        }
            //                        if (iteZ[1].Contains("(N)"))
            //                        {
            //                            valueRow["Average(N)"] = iteZ[1].Split(':')[1];

            //                        }
            //                    }
            //                }
            //                catch
            //                {

            //                }
            //                if (iteZ[1].Contains("Peak"))
            //                {
            //                    if (iteZ[1].Contains("(gf)"))
            //                    {
            //                        valueRow["Peak(Gf)"] = iteZ[1].Split(':')[1];
            //                    }
            //                    if (iteZ[1].Contains("(N)"))
            //                    {
            //                        valueRow["Peak(N)"] = iteZ[1].Split(':')[1];

            //                    }
            //                }
            //                dic[$"{type}-{iteZ[0]}"] = valueRow;

            //            }
            //        }
            //    }
            //}
            //foreach (DataRow item in dic.Values)
            //{
            //    item["ID"] = _SPEC[_key].Rows.Count + 1;
            //    _SPEC[_key].Rows.Add(item);
            //}
        }
        private DataTable setUpSpec()
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Id");
            dataTable.Columns.Add("ItemCode");
            dataTable.Columns.Add("LotNo");
            dataTable.Columns.Add("Type");
            dataTable.Columns.Add("Tape");
            dataTable.Columns.Add("Peak(N)");
            dataTable.Columns.Add("Average(N)");
            dataTable.Columns.Add("Peak(Gf)");
            dataTable.Columns.Add("Average(Gf)");
            return dataTable;
        }
        public PeelTestOnProductService(string iTEMCODE, string lOTNO, string type)
        {
            _ITEMCODE = iTEMCODE;
            _LOTNO = lOTNO;
            _TYPE = type;
            SetUpDT();
        }

        public PeelTestOnProductService()
        {
            SetUpDT();
        }

        public KeyValuePair<string, string> getItemCodeLotNo(string folderName)
        {
            string itemCode = "", lotNo = "";
            string[] s = folderName.Split('-', '_');
            if (s.Length > 4)
            {
                itemCode = s[1];
                try
                {
                    lotNo = ValidateService.lotNoHandle($"{s[2]}-{ValidateService.isDigit(s[3])}");
                }
                catch
                {
                    lotNo = ValidateService.lotNoHandle($"{s[2]}");
                }

            }
            else
            {

            }
            return new KeyValuePair<string, string>(itemCode, lotNo);
        }

        public void ReadData(string location, string pid)
        {
            _BEFOREIMAGE = new Dictionary<string, DataTable>();
            clearData();
            string msg = "";
            string makerName = FileFolderRepository.GetFolderName(location).Split('_')[1].Split('-')[0].Trim();
            KeyValuePair<string, string> pair = getItemCodeLotNo(location.Split('\\')[location.Split('\\').Count() - 1]);
            if (pair.Key == null || pair.Value == null)
            {
                throw new Exception("ItemCode lotno khong ton tai");
            }
            if (pair.Key.Trim().Equals(_ITEMCODE) && pair.Value.Trim().Equals(_LOTNO))
            {

            }
            else
            {
                _ITEMCODE = pair.Key;
                _LOTNO = pair.Value;
                msg += "Có thay đổi của itemCode lotNo\n";
            }

            bool primeSpec = false;
            try
            {
                MakerSpec(_ITEMCODE, _LOTNO);
                primeSpec = _SPEC[$"{_ITEMCODE} - {_LOTNO}"].Rows.Count <= 0;
            }
            catch
            {
                primeSpec = true;
                msg += "Spec chưa cài đặt!";
            }

            string[] sub = FileFolderRepository.GetSubFolders(location);
            foreach (string locationItem in sub)
            {
                string name = FileFolderRepository.GetFolderName(locationItem);
                switch (name)
                {
                    case "LINER":
                    case "PSA":
                        readFolderLinerPSA(locationItem, name, makerName, primeSpec);
                        break;
                    case "BF":
                    case "LINER TRUOC KEO":
                    case "ANH BF":
                        List<KeyValuePair<Image, string>> listIamge = FileFolderRepository.ListAllPictureInAFolder(locationItem, ".jpg");
                        listIamge.Sort(delegate (KeyValuePair<Image, string> item1, KeyValuePair<Image, string> item2)
                        {
                            if (int.TryParse(item1.Value.Split('.')[0], out int num1) && int.TryParse(item2.Value.Split('.')[0], out int num2))
                            {
                                return num1 > num2 ? 1 : -1;
                            }
                            else
                            {
                                return -1;

                            }
                        });
                        int i = 1;
                        string key = $"{_ITEMCODE} - {_LOTNO}";

                        if (_BEFOREIMAGE.TryGetValue(key, out DataTable dt))
                        {
                            foreach (KeyValuePair<Image, string> itemZ in listIamge)
                            {
                                DataRow row = dt.NewRow();
                                row["Id"] = i++;
                                row["ItemCode"] = _ITEMCODE;
                                row["LotNo"] = _LOTNO;
                                row["Picture"] = itemZ.Key;
                                dt.Rows.Add(row);
                            }
                            _BEFOREIMAGE[key] = dt;
                        }
                        else
                        {
                            dt = SetUpDT();
                            foreach (KeyValuePair<Image, string> itemZ in listIamge)
                            {
                                DataRow row = dt.NewRow();
                                row["Id"] = i++;
                                row["ItemCode"] = _ITEMCODE;
                                row["LotNo"] = _LOTNO;
                                row["Picture"] = itemZ.Key;
                                dt.Rows.Add(row);
                            }
                            _BEFOREIMAGE.Add(key, dt);
                        }
                        break;
                    default:
                        break;
                }
            }
            //fill PID 
            if (string.IsNullOrEmpty(pid))
            {
                msg += "Địa chỉ PRODUCTID không tồn tại!\n";
            }
            else
            {
                ProductIDService pidService = new ProductIDService(_ITEMCODE, _LOTNO, pid, new[] { "OQC", "peeling force" }, new[] { "Liner", "PSA" });
                foreach (string key in _DIC.Keys.ToArray())
                {
                    bool primePID = false;
                    // Debugger.Break();
                    string addValaue = "";
                    if (key.ToUpper().Contains("LINER"))
                    {
                        if (pidService._listFile.TryGetValue("Liner", out addValaue))
                        {
                            primePID = true;
                        }
                    }
                    if (key.ToUpper().Contains("PSA"))
                    {
                        if (pidService._listFile.TryGetValue("PSA", out addValaue))
                        {
                            primePID = true;
                        }
                    }
                    if (primePID)
                    {
                        List<string> listPID = pidService.getListProductID(addValaue);

                        foreach (DataRow row in _DIC[key].Rows)
                        {
                            if (listPID.Count >= int.Parse(row["Id"].ToString()))
                            {
                                row["ProductID"] = listPID[int.Parse(row["Id"].ToString()) - 1];
                            }
                        }
                    }


                }
            }
            CalculationSpec();
            throw new Exception(msg + "Lấy dữ liệu thành công!");
        }
        public void CalculationSpec()
        {
            // Hệ số quy đổi: 1 N = 101.9716 gf
            const double N_TO_GF = 101.9716;

            foreach (var kvp in _SPEC)
            {
                DataTable table = kvp.Value;
                if (table == null || table.Rows.Count == 0) continue;

                foreach (DataRow row in table.Rows)
                {
                    // Xử lý cặp cột Peak
                    ProcessColumnPair(row, "Peak(N)", "Peak(Gf)", N_TO_GF);

                    // Xử lý cặp cột Average
                    ProcessColumnPair(row, "Average(N)", "Average(Gf)", N_TO_GF);
                }
            }
            foreach(string key in _DIC.Keys)
            {
                CheckSpec(key);
            }
        }

        private void ProcessColumnPair(DataRow row, string colN, string colGf, double factor)
        {
            if (!row.Table.Columns.Contains(colN) || !row.Table.Columns.Contains(colGf))
                return;

            string valN = row[colN]?.ToString()?.Trim();
            string valGf = row[colGf]?.ToString()?.Trim();

            bool hasN = !string.IsNullOrEmpty(valN) && valN != DBNull.Value.ToString();
            bool hasGf = !string.IsNullOrEmpty(valGf) && valGf != DBNull.Value.ToString();

            // 1. Nếu cột N có dữ liệu mà cột Gf trống -> Quy đổi từ N sang Gf
            if (hasN && !hasGf)
            {
                row[colGf] = TransformRangeValue(valN, factor, true); // Tính và làm tròn Gf mới
                row[colN] = FormatOldValue(valN);                     // Làm tròn N cũ
            }
            // 2. Ngược lại: Nếu cột Gf có dữ liệu mà cột N trống -> Quy đổi từ Gf sang N
            else if (hasGf && !hasN)
            {
                row[colN] = TransformRangeValue(valGf, factor, false); // Tính và làm tròn N mới
                row[colGf] = FormatOldValue(valGf);                    // Làm tròn Gf cũ
            }
            // 3. Nếu cả 2 đều đã có dữ liệu -> Chỉ làm tròn lại cho đồng nhất (Tùy chọn)
            else if (hasN && hasGf)
            {
                row[colN] = FormatOldValue(valN);
                row[colGf] = FormatOldValue(valGf);
            }
        }

        // Hàm quy đổi đơn vị và làm tròn (cho giá trị mới)
        private string TransformRangeValue(string input, double factor, bool toGf)
        {
            string cleanInput = input.Replace("(", "").Replace(")", "");
            string[] parts = cleanInput.Split('~');

            if (parts.Length == 2 &&
                double.TryParse(parts[0].Trim(), out double minVal) &&
                double.TryParse(parts[1].Trim(), out double maxVal))
            {
                if (toGf)
                {
                    minVal = minVal * factor;
                    maxVal = maxVal * factor;
                }
                else
                {
                    minVal = minVal / factor;
                    maxVal = maxVal / factor;
                }

                return $"{minVal:F3} ~ {maxVal:F3}"; // Làm tròn 3 số
            }

            return input;
        }

        // Hàm CHỈ LÀM TRÒN lại chuỗi giá trị (cho giá trị cũ)
        private string FormatOldValue(string input)
        {
            string cleanInput = input.Replace("(", "").Replace(")", "");
            string[] parts = cleanInput.Split('~');

            if (parts.Length == 2 &&
                double.TryParse(parts[0].Trim(), out double minVal) &&
                double.TryParse(parts[1].Trim(), out double maxVal))
            {
                // Trả về đúng giá trị ban đầu nhưng ép làm tròn 3 số thập phân
                return $"{minVal:F3} ~ {maxVal:F3}";
            }

            return input;
        }
        public DataTable MakerTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Id");
            dt.Columns.Add("ProductID");
            dt.Columns.Add("ItemCode");
            dt.Columns.Add("LotNo");
            dt.Columns.Add("Picture", typeof(Image));
            dt.Columns.Add("Graph", typeof(Image));
            dt.Columns.Add("Peak(gf)");
            dt.Columns.Add("Average(gf)");
            dt.Columns.Add("Peak(N)");
            dt.Columns.Add("Average(N)");
            dt.Columns.Add("JudgementForce");
            dt.Columns.Add("JudgementMode");
            dt.Columns.Add("KeyDic");
            return dt;
        }
        private void clearData()
        {
            SetUpDT();
            _DIC = new Dictionary<string, DataTable>();
        }
        public string solveSigma(double[] list)
        {

            // 1. Tính giá trị trung bình (Mean / Average)
            double mean = list.Average();

            // 2. Tính độ lệch chuẩn mẫu (Sample Standard Deviation - STDEV.S)
            // Sử dụng n - 1 (độ tự do) theo chuẩn thống kê mẫu
            double variance = list.Sum(val => Math.Pow(val - mean, 2)) / (list.Length - 1);
            double stdDev = Math.Sqrt(variance);

            // 3. Tính giới hạn 4 Sigma
            double lowerLimit = mean - 4 * stdDev;
            double upperLimit = mean + 4 * stdDev;

            // Nếu giới hạn dưới không được âm (tùy thuộc vào đặc tính lực bóc/peeling force)
            lowerLimit = Math.Max(0, lowerLimit);

            return $"{lowerLimit} ~ {upperLimit}";
        }
        public DataTable readSubTapeFolder(string location, string keyDic, out string sigma )
        {
            DataTable dataTable = MakerTable();
            string locaImg = location;
            List<Double> listPeak = new List<double>();
            List<Double> listAverage = new List<double>();

            if (FileFolderRepository.GetSubFolders(location).Contains($"{location}\\ANH"))
            {
                locaImg += "\\ANH\\";
            }
            List<KeyValuePair<Image, string>> list = FileFolderRepository.ListAllPictureInAFolder(locaImg, ".jpg");
            List<string> files = FileFolderRepository.GetFileByExtension(location, "xlsx").ToList();
            list.Sort(delegate (KeyValuePair<Image, string> item1, KeyValuePair<Image, string> item2)
            {
                if (int.TryParse(item1.Value.Split('.')[0], out int i1) && int.TryParse(item2.Value.Split('.')[0], out int i2))
                {
                    if (i1 < i2)
                    {
                        return -1;
                    }
                    else
                    {
                        return 1;
                    }
                }
                else
                {
                    return -1;
                }
            });
            files.Sort(delegate (string item1, string item2)
            {
                string name1 = FileFolderRepository.GetFolderName(item1);
                string name2 = FileFolderRepository.GetFolderName(item2);

                // Tách số chính và số phụ bằng cách cắt chuỗi dựa vào dấu gạch ngang '-'
                // Ví dụ: "1-1" -> parts[0] = "1", parts[1] = "1"
                // Ví dụ: "1"   -> parts[0] = "1", không có phần 2
                string[] parts1 = name1.Split('.')[0].Split('-');
                string[] parts2 = name2.Split('.')[0].Split('-');

                bool parseMain1 = int.TryParse(parts1[0], out int main1);
                bool parseMain2 = int.TryParse(parts2[0], out int main2);

                if (parseMain1 && parseMain2)
                {
                    // So sánh số chính trước
                    if (main1 != main2)
                    {
                        return main1.CompareTo(main2);
                    }

                    // Nếu số chính bằng nhau, so sánh số phụ (nếu có)
                    int sub1 = (parts1.Length > 1 && int.TryParse(parts1[0], out int s1)) ? s1 : 0;
                    int sub2 = (parts2.Length > 1 && int.TryParse(parts2[1], out int s2)) ? s2 : 0; // Sửa lỗi cú pháp index của bạn thành s2

                    return sub1.CompareTo(sub2);
                }

                return string.Compare(item1, item2, StringComparison.Ordinal);
            });
            int i = 0;
            string peakRange = "";
            string averageRange = "";
            foreach (DataRow row in _SPEC[$"{_ITEMCODE} - {_LOTNO}"].Rows)
            {
                string rowA = $"{row["Tape"]} - {row["Type"]}";
                if (rowA.Equals($"{keyDic.Split('-', '_')[0]}- {keyDic.Split('-', '_')[2]}".Trim()))
                {
                    peakRange = row["Peak(N)"].ToString();
                    averageRange = row["Average(N)"].ToString();

                }
            }
            int z = 0;
            int step = 0;

            while (step < files.Count || i < list.Count)
            {
                DataRow row = dataTable.NewRow();
                row["ID"] = i + 1;
                row["ItemCode"] = _ITEMCODE;
                row["LotNo"] = _LOTNO;
                row["KeyDic"] = keyDic;
                if (z < list.Count)
                {
                    row["Picture"] = (Image)list[z].Key;
                    z++;
                }
                if (step < files.Count)
                {
                    KeyValuePair<Image, string> paire = solveFileData(files[step]);
                    int numZSS = int.Parse(FileFolderRepository.GetFileNameWithoutExtension(files[step]).Split('-')[0].Trim());
                    try
                    {

                        int iStep = int.Parse(FileFolderRepository.GetFileNameWithoutExtension(files[step + 1]).Split('-')[0].Trim());
                        if (numZSS  == iStep)
                        {
                            KeyValuePair<Image, string> paireZ = solveFileData(files[step + 1]);
                            paire = new KeyValuePair<Image, string>(paire.Key, $"{paireZ.Value.Split(',')[0]},{paire.Value.Split(',')[1]}");
                            step++;
                            
                        }
                        else
                        {
                            
                        }
                        
                        
                    }
                    catch
                    {

                    }
                    step++;
                    row["Graph"] = paire.Key;
                    if (keyDic.Contains("PSA"))
                    {
                        row["JudgementForce"] = "NG";
                        row["Peak(N)"] = paire.Value.Split(',')[0];
                        row["Average(N)"] = paire.Value.Split(',')[1];
                        bool prime1 = false, prime2 = false;

                        if (!string.IsNullOrEmpty(peakRange))
                        {
                            if (double.TryParse(row["Peak(N)"].ToString(), out double numZ) && double.TryParse(peakRange.Split('~')[0].Trim(), out double num1) && double.TryParse(peakRange.Split('~')[1].Trim(), out double num2) && num1 < numZ && numZ < num2)
                            {
                                prime1 = true;
                            }
                        }


                        if (!string.IsNullOrEmpty(averageRange))
                        {
                            if (double.TryParse(row["Average(N)"].ToString(), out double numZ) && double.TryParse(averageRange.Split('~')[0].Trim(), out double num1) && double.TryParse(averageRange.Split('~')[1].Trim(), out double num2) && num1 < numZ && numZ < num2)
                            {
                                prime2 = true;
                            }

                        }
                        if (prime1 && prime2)
                        {
                            row["JudgementForce"] = "OK";
                        }
                    }
                    else
                    {
                        row["Peak(gf)"] = paire.Value.Split(',')[0];
                        row["Average(gf)"] = paire.Value.Split(',')[1];
                        row["JudgementForce"] = "NG";
                        bool prime1 = false, prime2 = false;
                        if (double.TryParse(row["Peak(gf)"].ToString(), out double num))
                        {
                            row["Peak(N)"] = num * 0.0098;
                            if (!string.IsNullOrEmpty(peakRange))
                            {
                                if (double.TryParse(row["Peak(N)"].ToString(), out double numZ) && double.TryParse(peakRange.Split('~')[0].Trim(), out double num1) && double.TryParse(peakRange.Split('~')[1].Trim(), out double num2) && num1 < numZ && numZ < num2)
                                {
                                    prime1 = true;
                                }
                            }
                        }
                        if (double.TryParse(row["Average(gf)"].ToString(), out num))
                        {
                            row["Average(N)"] = num * 0.0098;
                            if (!string.IsNullOrEmpty(averageRange))
                            {
                                if (double.TryParse(row["Average(N)"].ToString(), out double numZ) && double.TryParse(averageRange.Split('~')[0].Trim(), out double num1) && double.TryParse(averageRange.Split('~')[1].Trim(), out double num2) && num1 < numZ && numZ < num2)
                                {
                                    prime2 = true;
                                }
                            }
                        }
                        if (prime1 && prime2)
                        {
                            row["JudgementForce"] = "OK";
                        }
                    }
                    listPeak.Add(double.Parse(row["Peak(N)"].ToString()));
                    listAverage.Add(double.Parse(row["Average(N)"].ToString()));
                }

                i++;
                dataTable.Rows.Add(row);

            }
            sigma = $"{solveSigma(listAverage.ToArray())}-{solveSigma(listPeak.ToArray())}";
            
            return dataTable;
        }
        private KeyValuePair<Image, string> solveFileData(string location)
        {
            KeyValuePair<Image, string> pair = new KeyValuePair<Image, string>();

            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            using (ExcelPackage package = new ExcelPackage(location))
            {
                using (ExcelWorksheet ws = package.Workbook.Worksheets[0])
                {
                    byte[] image = (ws.Drawings["Picture 1"] as ExcelPicture).Image.ImageBytes;
                    IDictionary<string, string> dic = ExportProcess.FindAddressByText(ws, new[] { "Max", "Average" }, true);
                    string value = ""; bool prime = false;
                    if (dic.TryGetValue("Max", out string address))
                    {
                        int i = 1;
                        while (i <= 20)
                        {
                            string addressZ = ws.Cells[ws.Cells[address].End.Row, ws.Cells[address].End.Column + i].Address;
                            var z = ws.Cells[addressZ].Value;
                            if (z != null)
                            {
                                value = z.ToString().Trim();
                                if (!string.IsNullOrEmpty(value))
                                {
                                    if (double.TryParse(value, out double ressz))
                                    {
                                        break;

                                    }
                                }
                            }
                            i++;
                        }
                        //if (string.IsNullOrEmpty(value))
                        //{
                        //    address = ws.Cells[ws.Cells[address].End.Row, ws.Cells[address].End.Column + 5].Address;
                        //    value += ws.Cells[address].Value.ToString().Trim();
                        //}
                    }
                    if (dic.TryGetValue("Average", out address))
                    {
                        int i = 1;
                        while (i <= 20)
                        {
                            string addressZ = ws.Cells[ws.Cells[address].End.Row, ws.Cells[address].End.Column + i].Address;
                            var z = ws.Cells[addressZ].Value;
                            if (z != null)
                            {
                                string valueZ = z.ToString().Trim();
                                if (!string.IsNullOrEmpty(valueZ))
                                {
                                    if (double.TryParse(valueZ, out double ressz))
                                    {
                                        value += "," + valueZ;
                                        break;

                                    }
                                }
                            }
                            i++;
                        }
                    }
                    pair = new KeyValuePair<Image, string>(TDMK_ImageConverter.ByteArrayToImage(image), value);
                }
            }
            return pair;
        }
        public void readFolderLinerPSA(string locationItem, string name, string makerName, bool primeSpec)
        {

            string[] subs = FileFolderRepository.GetSubFolders(locationItem);
            foreach (string s in subs)
            {
                string key = $"{FileFolderRepository.GetFolderName(s)} - {makerName}_{name}";
                _DIC.Add(key, readSubTapeFolder(s, key, out string sigma));

                if (primeSpec)
                {
                    DataRow row = _SPEC[$"{_ITEMCODE} - {_LOTNO}"].NewRow();
                    row["Id"] = _SPEC[$"{_ITEMCODE} - {_LOTNO}"].Rows.Count + 1;
                    row["ItemCode"] = _ITEMCODE;
                    row["LotNo"] = _LOTNO;
                    row["Type"] = name;
                    row["Tape"] = $"{FileFolderRepository.GetFolderName(s)}";
                    row["Average(N)"] = sigma.Split('-')[0];
                    row["Peak(N)"] = sigma.Split('-')[1];
                    _SPEC[$"{_ITEMCODE} - {_LOTNO}"].Rows.Add(row);
                }
                
            }
        }
        public bool checkNG()
        {
            foreach (string key in _DIC.Keys)
            {
                foreach (DataRow row in _DIC[key].Rows)
                {
                    if (!row["JudgementForce"].ToString().Contains("OK"))
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        //public DataTable _SPEC = new DataTable();
        //public Dictionary<string, DataTable> _DIC = new Dictionary<string, DataTable>();
        //public DataTable _BEFOREIMAGE;
        public DataTable getStructor()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("ProductID", typeof(string));
            dt.Columns.Add("ItemCode", typeof(string));
            dt.Columns.Add("LotNo", typeof(string));
            dt.Columns.Add("Graph", typeof(Image));
            dt.Columns.Add("Picture", typeof(Image));
            dt.Columns.Add("Peak(gf)", typeof(string));
            dt.Columns.Add("Average(gf)", typeof(string));
            dt.Columns.Add("Peak(N)", typeof(string));
            dt.Columns.Add("Average(N)", typeof(string));
            dt.Columns.Add("JudgementForce", typeof(string));
            dt.Columns.Add("JudgementMode", typeof(string));
            dt.Columns.Add("KeyDic", typeof(string));
            return dt;
        }
        public void SaveData(bool prime = false)
        {
            string fakeItemCode = $"{_ITEMCODE}" + (_TYPE == "Displacement" ? "&D" : "");
            DataTable dt = new DataTable();
            if (!prime)
            {
                dt = _dbContext.LoadDataTable("PT_ONPRODUCT_NAS", new[] { "ItemCode", "LotNo" }, new[] { fakeItemCode, _LOTNO });
                if (dt.Rows.Count > 0)
                {
                    throw new Exception("1402 - Đã có dữ liệu bạn muốn tiếp tục?");
                }
            }
            dt = getStructor();
            foreach (string key in _DIC.Keys)
            {
                if (!key.ToUpper().Contains("REFER"))
                {
                    foreach (DataRow rowZ in _DIC[key].Rows)
                    {
                        DataRow rowA = dt.NewRow();
                        foreach (DataColumn col in _DIC[key].Columns)
                        {
                            rowA[col.ColumnName] = rowZ[col.ColumnName];
                        }
                        dt.Rows.Add(rowA);
                    }
                }
            }
            NasRepository nas = new NasRepository();
            string location = nas.HandleImageDataTable(dt, "PT_ONPRODUCT", fakeItemCode, _LOTNO);
            string location_spec = nas.HandleImageDataTable(_SPEC[$"{_ITEMCODE} - {_LOTNO}"], "PT_ONPRODUCT_SEPC", fakeItemCode, _LOTNO);
            string location_before = nas.HandleImageDataTable(_BEFOREIMAGE[$"{_ITEMCODE} - {_LOTNO}"], "PT_ONPRODUCT_BEFOREIMAGE", fakeItemCode, _LOTNO);
            DataTable dataTable = _dbContext.GetTableStructure("PT_ONPRODUCT_NAS");
            DataRow row = dataTable.NewRow();
            row["ItemCode"] = fakeItemCode;
            row["LotNo"] = _LOTNO;
            row["LocationImage"] = location;
            row["LocationSpec"] = location_spec;
            row["LoactionBefore"] = location_before;
            row["Data"] = ConverterService.DataTableToJson(dt);
            row["DataSpec"] = ConverterService.DataTableToJson(_SPEC[$"{_ITEMCODE} - {_LOTNO}"]);
            row["DataBefore"] = ConverterService.DataTableToJson(_BEFOREIMAGE[$"{_ITEMCODE} - {_LOTNO}"]);
            dataTable.Rows.Add(row);
            int res = _dbContext.BuckDataTable(dataTable, "PT_ONPRODUCT_NAS", new[] { "ItemCode", "LotNo" }, null, "ID");
            //int res = _dbContext.BuckDataTable(dt, "PT_ONPRODUCT", new[] { "ItemCode", "LotNo" }, null, "Id");
            //res += _dbContext.BuckDataTable(_SPEC[$"{_ITEMCODE} - {_LOTNO}"], "PT_ONPRODUCT_SEPC", new[] { "ItemCode", "LotNo" }, null, "Id");
            //res += _dbContext.BuckDataTable(_BEFOREIMAGE[$"{_ITEMCODE} - {_LOTNO}"], "PT_ONPRODUCT_BEFOREIMAGE", new[] { "ItemCode", "LotNo" }, null, "Id");



            DataTable dtZ = _dbContext.LoadDataTable("AIR_BUBBLE_REFER", new[] { "ItemCodeRefer", "LotNoRefer" }, new[] { fakeItemCode, _LOTNO });
            DataTable dtZ2 = _dbContext.LoadDataTable("AIR_BUBBLE_REFER", new[] { "ItemMain", "LotMain" }, new[] { fakeItemCode, _LOTNO });
            if (dtZ.Rows.Count <= 0 && dtZ2.Rows.Count <= 0)
            {
                DataTable dtRes = _dbContext.GetTableStructure("AIR_BUBBLE_AVAILABLE");
                DataRow rowZ = dtRes.NewRow();
                rowZ["ItemCode"] = fakeItemCode;
                rowZ["LotNo"] = _LOTNO;
                rowZ["Status"] = 1;
                dtRes.Rows.Add(rowZ);
                res += _dbContext.BuckDataTable(dtRes, "AIR_BUBBLE_AVAILABLE", new[] { "ItemCode", "LotNo" }, null, "ID");
            }

            throw new Exception($"{res} lưu thành công!");
        }

        public void LoadData(bool legacy, bool refer = false)
        {
            string msg = "";
            string ITEMCODE = _ITEMCODE.Trim();
            string LOTNO = _LOTNO.Trim();
            string fakeItemCode = $"{_ITEMCODE}" + (_TYPE == "Displacement" ? "&D" : "");
            if (string.IsNullOrEmpty(fakeItemCode) || string.IsNullOrEmpty(_LOTNO))
            {
                throw new Exception("ItemCode hoặc LotNo không được để trống!");
            }
            List<string> list = new List<string>();
            list.Add($"{fakeItemCode} - {LOTNO}");
            if (refer)
            {
                DataTable dtZ = _dbContext.LoadDataTable("AIR_BUBBLE_REFER", new[] { "ItemMain", "LotMain" }, new[] { fakeItemCode, _LOTNO });
                if (dtZ.Rows.Count <= 0)
                {
                    msg += "Không có refer!";
                }
                else
                {

                    foreach (DataRow row in dtZ.Rows)
                    {
                        list.Add($"{row["ItemCodeRefer"]}-{row["LotNoRefer"]}");
                    }
                }
            }

            _DIC = new Dictionary<string, DataTable>();
            foreach (string item in list)
            {
                if (string.IsNullOrEmpty(item))
                {
                    break;
                }
                ITEMCODE = item.Split('-')[0].Trim();
                LOTNO = item.Split('-')[1].Trim();
                string keyZ = $"{ITEMCODE.Replace("&D", "")} - {LOTNO}";
                DataTable dataTable = new DataTable();
                if (legacy)
                {

                    _BEFOREIMAGE.Add(keyZ, _dbContext.LoadDataTable("PT_ONPRODUCT_BEFOREIMAGE", new[] { "ItemCode", "LotNo" }, new[] { ITEMCODE, LOTNO }));
                    _SPEC.Add(keyZ, _dbContext.LoadDataTable("PT_ONPRODUCT_SEPC", new[] { "ItemCode", "LotNo" }, new[] { ITEMCODE, LOTNO }));
                    dataTable = _dbContext.LoadDataTable("PT_ONPRODUCT", new[] { "ItemCode", "LotNo" }, new[] { ITEMCODE, LOTNO });
                }
                else
                {
                    dataTable = _dbContext.LoadDataTable("PT_ONPRODUCT_NAS", new[] { "ItemCode", "LotNo" }, new[] { ITEMCODE, LOTNO });
                    if (dataTable.Rows.Count <= 0)
                    {
                        if (!refer)
                        {
                            throw new Exception("No Data");
                        }
                        else
                        {
                            return;
                        }
                    }
                    DataRow rowZ = dataTable.Rows[0];
                    NasRepository _nas = new NasRepository();
                    DataTable beforeDT = ConverterService.JsonToDataTable(rowZ["DataBefore"].ToString());
                    _nas.MergeDataTable(beforeDT, "PT_ONPRODUCT_BEFOREIMAGE", ITEMCODE, LOTNO, rowZ["LoactionBefore"].ToString());
                    _BEFOREIMAGE.Add(keyZ, beforeDT);
                    DataTable specDT = ConverterService.JsonToDataTable(rowZ["DataSpec"].ToString());
                    _nas.MergeDataTable(specDT, "PT_ONPRODUCT_SEPC", ITEMCODE, LOTNO, rowZ["LocationSpec"].ToString());
                    _SPEC.Add(keyZ, specDT);
                    dataTable = ConverterService.JsonToDataTable(rowZ["Data"].ToString());
                    _nas.MergeDataTable(dataTable, "PT_ONPRODUCT_NAS", ITEMCODE, LOTNO, rowZ["LocationImage"].ToString());

                }
                //_DIC
                foreach (DataRow row in dataTable.Rows)
                {
                    string key = row["KeyDic"].ToString() + (refer && !ITEMCODE.Equals(fakeItemCode) || !LOTNO.Equals(_LOTNO) ? " REFER " + $"({fakeItemCode}-{LOTNO})" : "");
                    if (_DIC.TryGetValue(key, out DataTable dt))
                    {
                        DataRow newRow = dt.NewRow();
                        foreach (DataColumn col in dataTable.Columns)
                        {
                            newRow[col.ColumnName] = row[col.ColumnName];
                        }
                        dt.Rows.Add(newRow);
                        _DIC[key] = dt;
                    }
                    else
                    {
                        dt = dataTable.Clone();
                        DataRow newRow = dt.NewRow();
                        foreach (DataColumn col in dataTable.Columns)
                        {
                            newRow[col.ColumnName] = row[col.ColumnName];
                        }
                        dt.Rows.Add(newRow);
                        _DIC.Add(key, dt);
                    }
                }
            }

        }

        public void Export(bool legacy)
        {
            string msg = "";
            if (_DIC == null || _DIC.Count <= 0)
            {
                try
                {
                    LoadData(legacy);
                }
                catch
                {

                }
            }
            ExportProcess exportProcess = new ExportProcess();

            using (ExcelPackage package = exportProcess.FindFormatProcess("AIR", _ITEMCODE, _LOTNO))
            {
                using (ExcelWorksheet ws = exportProcess.FindSheet(package, "Liner peel test (On product)"))
                {
                    SetUpExport(ws, true);
                    FillData(ws, true);
                    using (ExcelWorksheet ws2 = exportProcess.FindSheet(package, "PSA peel test (On product)"))
                    {

                        SetUpExport(ws2);
                        FillData(ws2, false);
                        exportProcess.SaveExcelWorksheet(package, "Liner peel test (On product):PSA peel test (On product)", $"{_ITEMCODE.Trim()}-{_LOTNO.Trim()}");

                    }
                }
            }
        }

        private void FillData(ExcelWorksheet ws, bool liner)
        {
            List<string> listHelder = new[] { "supplier" }.ToList();
            string stage = "";
            if (liner)
            {
                stage = "Liner";
                listHelder.Add("Liner peeling");
            }
            else
            {
                stage = "PSA";
                listHelder.Add("PSA peeling");
            }
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(ws, listHelder.ToArray());
            int rCom = ws.Cells[dic[stage + " peeling"].Split('-')[0]].End.Row;
            foreach (var item in dic["supplier"].Split('-'))
            {
                int r = ws.Cells[item].End.Row;
                if (rCom < r)
                {
                    string valueZ = ws.Cells[item].Value.ToString();
                    string tape;
                    if(_TYPE == "Displacement")
                    {
                        tape = valueZ.Split('_')[2].Trim();
                    }
                    else
                    {
                        tape = valueZ.Split('(')[1].Replace(")", "").Trim();
                    }
                    string keyMAIN = "";
                    if (string.IsNullOrEmpty(valueZ.Split('_', '(')[1].ToUpper().Replace("SUPPLIER", "").Replace("NAME", "").Trim()))
                    {
                        //peek key no refer
                        foreach (string key in _DIC.Keys)
                        {
                            if (!key.Contains("REFER") && key.ToUpper().Split('-')[0].Trim().Equals(tape.ToUpper()) && key.Split('_')[1].Trim().ToUpper().Equals(stage.ToUpper()))
                            {
                                // maker name
                                ws.Cells[item].Value = $"{valueZ.Split('_')[0]}_ {key.Split('-', '_')[1]} {valueZ.Split('_')[2]}";
                                keyMAIN = key;
                                break;
                            }
                        }
                    }
                    else
                    {
                        foreach (string key in _DIC.Keys)
                        {
                            string z = key.Split('-', '_')[1].Trim();
                            string za = valueZ.Split('_', '(')[1].ToUpper().Replace("SUPPLIER", "").Replace("NAME", "").Trim().ToUpper();
                            if (z.Equals(za) && key.Contains("REFER") && key.Split('-')[0].Trim().Equals(tape) && key.Split('_')[1].Trim().ToUpper().Contains(stage.ToUpper()))
                            {
                                // peek key co refer
                                keyMAIN = key;
                                break;
                            }
                        }
                    }
                    int i = 0;
                    string icln = keyMAIN.Contains("REFER") ? keyMAIN.Split('(')[1].Replace(")", "").Trim() : $"{_ITEMCODE}-{_LOTNO}";
                    icln = icln.Replace("-", " - ");
                    double peakMin = 0, peakMax = 0, averageMin = 0, averageMax = 0;
                    double peakGFMin = 0, peakGFMax = 0, averageGFMin = 0, averageGFMax = 0;
                    if (_BEFOREIMAGE.TryGetValue(icln, out DataTable valueDT))
                    {
                        if (_SPEC.TryGetValue(icln, out DataTable spec))
                        {
                            foreach (DataRow row in spec.Rows)
                            {
                                if (row["Tape"].ToString().Contains(tape) && row["Type"].ToString().ToUpper().Contains(stage.ToUpper()))
                                {
                                    //Debugger.Break();
                                    try
                                    {

                                        string peak = row["Peak(Gf)"].ToString();
                                        if (peak.Contains('~'))
                                        {

                                            peakGFMin = double.Parse(peak.Split('~')[0]);
                                            peakGFMax = double.Parse(peak.Split('~')[1]);
                                        }
                                    }
                                    catch
                                    {

                                    }

                                    //Debugger.Break();
                                    try
                                    {

                                        string peak = row["Peak(N)"].ToString();
                                        if (peak.Contains('~'))
                                        {

                                            peakMin = double.Parse(peak.Split('~')[0]);
                                            peakMax = double.Parse(peak.Split('~')[1]);
                                        }
                                    }
                                    catch
                                    {

                                    }
                                    try
                                    {

                                        string average = row["Average(N)"].ToString();
                                        if (average.Contains('~'))
                                        {
                                            averageMin = double.Parse(average.Split('~')[0]);
                                            averageMax = double.Parse(average.Split('~')[1]);
                                        }
                                    }
                                    catch
                                    {

                                    }
                                    try
                                    {

                                        string average = row["Average(Gf)"].ToString();
                                        if (average.Contains('~'))
                                        {
                                            averageGFMin = double.Parse(average.Split('~')[0]);
                                            averageGFMax = double.Parse(average.Split('~')[1]);
                                        }
                                    }
                                    catch
                                    {

                                    }
                                    break;
                                }
                            }
                        }
                        foreach (DataRow row in _DIC[keyMAIN].Rows)
                        {
                            string address = ExportProcess.AddColumn(item, 2 + i);
                            ws.Cells[ExportProcess.AddRow(address, -1)].Value = row["ProductID"];
                            if (liner)
                            {
                                address = ExportProcess.AddRow(address, 1);
                                if (valueDT.Rows.Count > i)
                                {
                                    ExportProcess.InsertImageToCell(ws, ws.Cells[address], (byte[])valueDT.Rows[i]["Picture"], $"{Guid.NewGuid()}");
                                }

                            }
                            ExportProcess.InsertImageToCell(ws, ws.Cells[ExportProcess.AddRow(address, 1)], (byte[])row["Picture"], $"{Guid.NewGuid()}");
                            ExportProcess.InsertImageToCell(ws, ws.Cells[ExportProcess.AddRow(address, 2)], (byte[])row["Graph"], $"{Guid.NewGuid()}");
                            if (liner)
                            {
                                address = ExportProcess.AddRow(address, 1);
                                if (double.TryParse(row["Peak(gf)"].ToString(), out double res))
                                {
                                    ws.Cells[ExportProcess.AddRow(address, 2)].Value = Math.Round(res, 2);
                                }
                                address = ExportProcess.AddRow(address, 1);
                                if (double.TryParse(row["Average(gf)"].ToString(), out res))
                                {
                                    ws.Cells[ExportProcess.AddRow(address, 2)].Value = Math.Round(res, 2);
                                }
                            }
                            if (double.TryParse(row["Peak(N)"].ToString(), out double resZ))
                            {
                                ws.Cells[ExportProcess.AddRow(address, 3)].Value = Math.Round(resZ, 2);
                            }
                            if (double.TryParse(row["Average(N)"].ToString(), out resZ))
                            {
                                ws.Cells[ExportProcess.AddRow(address, 4)].Value = Math.Round(resZ, 2);
                            }

                            //_SPEC[keyMAIN].Rows
                            ws.Cells[ExportProcess.AddRow(address, 5)].FormulaR1C1 = $"=IF(AND(R[-1]C<{averageMax},R[-1]C>{averageMin},R[-2]C<{peakMax},R[-2]C>{peakMin}),\"OK\",\"NG\")";
                            ws.Cells[ExportProcess.AddRow(address, 6)].Value = row["JudgementMode"];
                            // Debugger.Break();
                            i++;
                        }

                    }
                    if (!liner)
                    {
                        string addressZ = ExportProcess.AddRow(ExportProcess.AddColumn(item, 2), 10);
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 0), -1)].Value = $"{peakMin} ~ {peakMax}";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 1), -1)].Value = $"{averageMin} ~ {averageMax}";

                        ws.Cells[addressZ].FormulaR1C1 = $"=MIN(R[-7]C:R[-7]C[31])";
                        ws.Cells[ExportProcess.AddColumn(addressZ, 1)].FormulaR1C1 = $"=MIN(R[-6]C[-1]:R[-6]C[30])";
                        ws.Cells[ExportProcess.AddRow(addressZ, 1)].FormulaR1C1 = $"=MAX(R[-8]C:R[-8]C[31])";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 1), 1)].FormulaR1C1 = $"=MAX(R[-7]C[-1]:R[-7]C[30])";
                        ws.Cells[ExportProcess.AddRow(addressZ, 2)].FormulaR1C1 = $"=Average(R[-9]C:R[-9]C[31])";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 1), 2)].FormulaR1C1 = $"=Average(R[-8]C[-1]:R[-8]C[30])";
                        ws.Cells[ExportProcess.AddRow(addressZ, 3)].FormulaR1C1 = $"=STDEV(R[-10]C:R[-10]C[31])";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 1), 3)].FormulaR1C1 = $"=STDEV(R[-9]C[-1]:R[-9]C[30])";
                        ws.Cells[ExportProcess.AddRow(addressZ, 4)].FormulaR1C1 = $"=MIN((R[-2]C - {peakMin})/(3 * R[-1]C), ({peakMax} - R[-2]C)/(3 * R[-1]C))";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 1), 4)].FormulaR1C1 = $"=MIN((R[-2]C - {averageMin})/(3 * R[-1]C),({averageMax} - R[-2]C)/(3 * R[-1]C))";
                    }
                    else
                    {
                        string addressZ = ExportProcess.AddRow(ExportProcess.AddColumn(item, 2), 13);
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 0), -1)].Value = $"{peakGFMin} ~ {peakGFMax}";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 1), -1)].Value = $"{peakMin} ~ {peakMax}";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 2), -1)].Value = $"{averageGFMin} ~ {averageGFMax}";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 3), -1)].Value = $"{averageMin} ~ {averageMax}";

                        ws.Cells[addressZ].FormulaR1C1 = $"=MIN(R[-9]C:R[-9]C[31])";
                        ws.Cells[ExportProcess.AddColumn(addressZ, 1)].FormulaR1C1 = $"=MIN(R[-7]C[-1]:R[-7]C[31])";
                        ws.Cells[ExportProcess.AddColumn(addressZ, 2)].FormulaR1C1 = $"=MIN(R[-8]C[-2]:R[-8]C[29])";
                        ws.Cells[ExportProcess.AddColumn(addressZ, 3)].FormulaR1C1 = $"=MIN(R[-6]C[-3]:R[-6]C[28])";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 0), 1)].FormulaR1C1 = $"=MAX(R[-10]C:R[-10]C[31])";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 1), 1)].FormulaR1C1 = $"=MAX(R[-8]C[-1]:R[-8]C[30])";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 2), 1)].FormulaR1C1 = $"=MAX(R[-9]C[-2]:R[-9]C[29])";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 3), 1)].FormulaR1C1 = $"=MAX(R[-7]C[-3]:R[-7]C[28])";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 0), 2)].FormulaR1C1 = $"=AVERAGE(R[-11]C:R[-11]C[31])";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 1), 2)].FormulaR1C1 = $"=AVERAGE(R[-9]C[-1]:R[-9]C[30])";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 2), 2)].FormulaR1C1 = $"=AVERAGE(R[-10]C[-2]:R[-10]C[29])";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 3), 2)].FormulaR1C1 = $"=AVERAGE(R[-8]C[-3]:R[-8]C[28])";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 0), 3)].FormulaR1C1 = $"=STDEV(R[-12]C:R[-12]C[31])";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 1), 3)].FormulaR1C1 = $"=STDEV(R[-10]C[-1]:R[-10]C[30])";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 2), 3)].FormulaR1C1 = $"=STDEV(R[-11]C[-2]:R[-11]C[29])";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 3), 3)].FormulaR1C1 = $"=STDEV(R[-9]C[-3]:R[-9]C[28])";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 0), 4)].FormulaR1C1 = $"=MIN((R[-2]C - {peakGFMin} )/(3 * R[-1]C), ( {peakGFMax} - R[-2]C)/(3 * R[-1]C))";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 1), 4)].FormulaR1C1 = $"=MIN((R[-2]C - {peakMin})/(3 * R[-1]C), ({peakMax} - R[-2]C)/(3 * R[-1]C))";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 2), 4)].FormulaR1C1 = $"=MIN((R[-2]C - {averageGFMin})/(3 * R[-1]C), ({averageGFMax} - R[-2]C)/(3 * R[-1]C))";
                        ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(addressZ, 3), 4)].FormulaR1C1 = $"=MIN((R[-2]C - {averageMin})/(3 * R[-1]C), ({averageMax} - R[-2]C)/(3 * R[-1]C))";

                    }
                }
            }

        }

        public void SetUpExport(ExcelWorksheet ws, bool liner = false)
        {
            List<string> list = new[] { "Tape", "Flex SN", "Sample 32", "CPK" }.ToList();
            string caseZ = "";
            if (liner)
            {
                caseZ = "Liner peeling";
            }
            else
            {
                caseZ = "PSA peeling";

            }
            list.Add(caseZ);
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(ws, list.ToArray());

            int addBase = ws.Cells[dic[caseZ].Split('-')[0]].End.Row;
            List<string> ne = new List<string>();
            foreach (var item in dic["Tape"].Split('-'))
            {
                int r = ws.Cells[item].End.Row;
                if (r > addBase) 
                {
                    ne.Add(item);
                }
            }
            dic["Tape"] = string.Join("-", ne.ToArray());
            string[] keys = _DIC.Keys.ToArray();
            Dictionary<string, string> listZ = new Dictionary<string, string>();
            foreach (string key in keys)
            {
                //if (_DIC[key].Rows.Count > 0)
                //{
                if (key.Contains(liner ? "LINER" : "PSA"))
                {
                    string tapeZ = key.Split('-')[0].Trim().ToUpper();
                    if (listZ.TryGetValue(tapeZ, out string value))
                    {
                        value += "," + key;
                        listZ[tapeZ] = value;
                    }
                    else
                    {

                        listZ.Add(tapeZ, key);
                    }
                }
                //}

            }
            keys = dic["Tape"].Split('-');

            for (int i = keys.Length - 1; i >= 0; i--)
            {
                //Debugger.Break();
                string tape;
                if(_TYPE == "Displacement")
                {
                    tape = ws.Cells[keys[i]].Value.ToString().Split('_')[2].Trim();
                }
                else
                {
                    tape = ws.Cells[keys[i]].Value.ToString().Split('(', ')')[1].Trim();
                }
                int z = listZ[tape.ToUpper()].Split(',').Count() - 1;
                for (int iZ = 0; iZ < z; iZ++)
                {
                    string tapeAdd = dic["Tape"].Split('-')[i];
                    string dist = ExportProcess.AddRow(ExportProcess.AddColumn(dic["CPK"].Split('-')[i], -1), 1);
                    string flexSN = dic["Flex SN"].Split('-')[i];
                    string sample = dic["Sample 32"].Split('-')[i];
                    string range = $"{flexSN}:{ws.Cells[ws.Cells[dist].End.Row, ws.Cells[sample].End.Column]}";
                    int r = ws.Cells[tapeAdd].End.Row - ws.Cells[flexSN].End.Row;
                    int c = ws.Cells[tapeAdd].End.Column - ws.Cells[flexSN].End.Column;

                    tapeAdd = ws.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(dist, c), r + 1)].Address;
                    new ExportProcess().CopyAndInsert(ws, range, ref dist, true);


                    string valueTape = ws.Cells[tapeAdd].Value.ToString();
                    ws.Cells[tapeAdd].Value = $"{valueTape.Split('_')[0]}_{listZ[tape].Split(',')[iZ + 1].Split('-', '_')[1].Trim()} {valueTape.Split('_')[1]}";

                }

            }
        }
    }
}
