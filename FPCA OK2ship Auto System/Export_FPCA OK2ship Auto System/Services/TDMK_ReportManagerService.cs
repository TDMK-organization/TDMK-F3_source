using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Export_FPCA_OK2ship_Auto_System.Repositories;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DataTable = System.Data.DataTable;
using Sheets = DocumentFormat.OpenXml.Spreadsheet.Sheets;
using Workbook = DocumentFormat.OpenXml.Spreadsheet.Workbook;

namespace Export_FPCA_OK2ship_Auto_System.Services
{
    public class TDMK_ReportManagerService
    {
        private string _DBNAME = "Reports";

        public static string[] categories = new[]
        {
            "Coverpage",
            "Rev History",
            "User Guideline",
            "Low CPK Action",
            "Declaration",
            "Table of Contents",
            "Deviation",
            "Assy Yield",
            "FAI",
            "Dimension without parentheses",
            "Dimension with parentheses",
            "OQC Test",
            "Cross Section",
            "GAP Connector",
            "Peel Test",
            "Peel Test without SUS",
            "Pull Test",
            "(IQC Unmating) Pull Test",
            "Shear test",
            "OQC B2B Mating-Unmating",
            "IQC Liner peeling (Coupon)",
            "IQC PSA peeling (Coupon)",
            "Liner peel test (On product)",
            "PSA peel test (On product)",
            "Air bubble btw Liner-PSA",
            "Air bubble btw PSA-FPC",
            "ORT-Assy ",
            "Flex bending",
            "Flex bending - ICT data",
            "Flex bending - X - Ray picture",
            "Thermal Cycling & bending",
            "TC & bending - ICT data ",
            "TC & bending - X - Ray picture",
            "Heat soak & bending",
            "HS & bending - ICT data",
            "HS & bending - X - Ray picture",
            "Thermal Cycling",
            "Thermal Cycling - ICT data",
            "Heat Soak and Recovery",
            "Heat soak & Recovery - ICT data",
            "Thermal Shock",
            "Thermal Shock - ICT data",
            "Environment Test",
            "Impedance", "Switch Quality",
            "ACF",
            "SEM BSE & Binarization",
            "Bar Code Verification",
            "Packaging",
            "Mishandling test",
            "Process flow",
            "Config comparison"
        };

        public string __Build { get; set; }
        public string __ItemName { get; set; }
        public string __Config { get; set; }

        /// <summary>
        /// Trạng thái của từng mục trong báo cáo 
        /// </summary>
        public Dictionary<string, string> __STATUS = new Dictionary<string, string>();

        public List<string> listcategory = new List<string>();

        public TDMK_ReportManagerService(string itemName, string build, string config)
        {
            itemName = itemName.Trim();
            build = build.Trim();
            config = config.Trim();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="progress"></param>
        /// <param name="reportForm">true is get sheet from report, false get sheet from format</param>
        public async Task ExportAll(List<string> _listCategory, IProgress<float> progress = null,
            bool reportForm = false, bool dbFirst = false)
        {
            if (_listCategory.Count <= 0)
            {
                _listCategory = listcategory;
            }

            string destinationFilePath = new ExportProcess().getExportLocation();
            string targetFolder = Path.Combine(destinationFilePath,
                $"{__ItemName}_{__Build}_{__ItemName}");
            List<string> allFile = new List<string>();
            float num = 80 / _listCategory.Count;
            float percent = 0;

            // [Step 1]: Lấy dữ liệu lẻ 
            foreach (string item in _listCategory)
            {
                progress.Report(percent);
                string status = "";
                try
                {
                    if (__STATUS.TryGetValue(item, out string value))
                    {
                        //Debugger.Break();
                        // trường hợp nếu là export từ nas
                        if (value.ToUpper().Equals("NAS"))
                        {
                            status = ExportCategoryNas(item, targetFolder);
                        }
                        else if (value.ToUpper().Equals("NONE"))
                        {
                            // bỏ qua
                        }
                        else if (value.ToUpper().Contains("OK2SHIP") || value.ToUpper().Contains("OK2BUILD"))
                        {
                            Dictionary<string, string> dic = new Dictionary<string, string>();
                            string[] valueArr = value.ToString().Split('(', ')', '-');
                            try
                            {
                                dic.Add("TargetFolder", targetFolder);
                                dic.Add("ItemCode", valueArr[1]);
                                dic.Add("LotNo", valueArr[2]);
                                dic.Add("Maker", valueArr[3]);
                            }
                            catch
                            {
                            }

                            status = ExportCategoryDB(item, item, reportForm, dbFirst, dic);
                        }
                        else
                        {
                            // Trường hợp đặc biệt
                            Debugger.Break();
                            throw new Exception("Không có trường hợp này!");
                        }
                    }
                }
                catch (Exception ex)
                {
                    status = ex.Message;
                }

                allFile.Add($"{item}:{status}");
                progress.Report(percent);
                percent += num;
            }

            // [Step 2]: Gộp file
            progress.Report(80);
            // try
            // {
            //     // 2.1. Lấy đường dẫn động tới thư mục chứa phần mềm
            //     string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            //
            //     // Tạo đường dẫn chuẩn đến file .exe
            //     string exePath = Path.Combine(baseDirectory, "VHX-IMADA", "MergeService", "MergeHeavyExcelService.exe");
            //
            //     // Kiểm tra xem file .exe có tồn tại không
            //     if (!File.Exists(exePath))
            //     {
            //         throw new Exception("Không tìm thấy file dịch vụ MergeHeavyExcelService.exe tại: " + exePath);
            //     }
            //
            //     // 2. Tạo chuỗi arguments (tham số)
            //     StringBuilder arguments = new StringBuilder();
            //
            //     // Nối danh sách các file input (Mỗi đường dẫn bọc trong dấu ngoặc kép "")
            //     foreach (string file in allFile)
            //     {
            //         arguments.Append($"\"{file}\" ");
            //     }
            //
            //     string dirOutPut = new ExportProcess().getLocationSave();
            //     // Nối file output vào cuối cùng (Bọc trong ngoặc kép "")
            //     arguments.Append($"\"{dirOutPut}\"");
            //
            //     // 3. Cấu hình tiến trình (Process) chạy file exe
            //     ProcessStartInfo startInfo = new ProcessStartInfo
            //     {
            //         FileName = exePath,
            //         Arguments = arguments.ToString(),
            //         UseShellExecute = false,
            //         CreateNoWindow = true,
            //         WindowStyle = ProcessWindowStyle.Hidden,
            //         // THÊM: Đọc luồng để bắt lỗi từ Python
            //         RedirectStandardOutput = true,
            //         RedirectStandardError = true,
            //         StandardOutputEncoding = Encoding.UTF8,
            //         StandardErrorEncoding = Encoding.UTF8
            //     };
            //
            //     // 4. Khởi chạy tiến trình
            //     using (Process process = new Process())
            //     {
            //         process.StartInfo = startInfo;
            //         process.Start();
            //
            //         // Đọc lại toàn bộ nội dung mà Python đã print() ra
            //         string outputLog = process.StandardOutput.ReadToEnd();
            //         string errorLog = process.StandardError.ReadToEnd();
            //
            //         process.WaitForExit();
            //
            //         // 5. Kiểm tra mã trả về
            //         if (process.ExitCode != 0)
            //         {
            //             // Bắn ra Exception kèm theo log của file Python để dễ debug
            //             throw new Exception(
            //                 $"Dịch vụ ghép file thất bại (Mã lỗi: {process.ExitCode}).\n\nLog hệ thống:\n{outputLog}\n{errorLog}");
            //         }
            //     }
            //
            //     // 6. MỞ FILE SAU KHI GHÉP THÀNH CÔNG
            //     Process.Start(new ProcessStartInfo(dirOutPut)
            //     {
            //         UseShellExecute = true
            //     });
            // }
            // catch (Exception ex)
            // {
            //     throw new Exception($"Quá trình mergefile lỗi:{ex.Message}");
            // }

            progress.Report(100);
        }


        public string ExportCategoryNas(string nameReal, string targetFolder)
        {
            IList<string> listFile = FileFolderRepository.ListAllFileInFolder(targetFolder, "xlsx");
            foreach (string item in listFile)
            {
                string file = FileFolderRepository.GetFileName(item).ToUpper().Replace(".XLSX", "");
                if (file.Contains(nameReal.ToUpper()))
                {
                    return item;
                }
            }

            return "Không có dữ liệu nas";
        }

        /// <summary>
        //
        /// </summary>
        /// <param name="nameReal">Tên thực tế của hạng mục (sau khi export sẽ lưu lại dưới tên này</param>
        /// <param name="category">Tên mục sẽ export (dựa vào tên này để chọn quy trình để xuất)</param>
        /// <param name="reportForm">Lấy báo cáo từ true: từ báo cáo gốc, false: từ</param>
        /// <param name="dbFirst">true:ưu tiên xuất dữ liệu từ database\ false kiểm tra xem có sẵn báo cáo đấy chưa nếu có thì sử dụng </param>
        public string ExportCategoryDB(string nameReal, string category, bool reportForm, bool dbFirst,
            Dictionary<string, string> info)
        {
            //Debugger.Break();

            // Nếu ưu tiên sử dụng nas
            if (info.TryGetValue("TargetFolder", out string targetFolder))
            {
                IList<string> listFile = FileFolderRepository.ListAllFileInFolder(targetFolder, "xlsx");
                foreach (string item in listFile)
                {
                    string file = FileFolderRepository.GetFileName(item).ToUpper().Replace(".xlsx", "");
                    if (file.Equals(nameReal))
                    {
                        if (!dbFirst)
                        {
                            return item;
                        }
                        else
                        {
                            targetFolder = item;
                            break;
                        }
                    }
                }
            }

            if (!info.TryGetValue("ItemCode", out string itemCode))
            {
                throw new Exception("ItemCode không tồn tại trong thông tin xuất!");
            }

            //Lấy format 
            using (ExcelPackage package = reportForm
                       ? new ExportProcess().FindFormatWithItemCode(itemCode)
                       : new ExportProcess().OpenFileExcel(targetFolder))
            {
                // cho các thụuật toán xuất vào đây
                Debugger.Break();
                switch (category)
                {
                    case "Cross Section":
                        if (info.TryGetValue("LotNo", out string lotNo))
                        {
                           return CrossSectionService.ExportAll(package, itemCode, lotNo);
                            
                        }
                        else
                        {
                            
                        }

                        break;
                }
            }


            return "";
        }

        public void initData()
        {
            __STATUS = new Dictionary<string, string>();
            listcategory = new List<string>();
            foreach (string item in categories)
            {
                __STATUS.Add(item, "None");
                listcategory.Add(item);
            }
        }

        /// <summary>
        /// Lưu dữ liệu
        /// </summary>
        public void Save(bool checkDB = true)
        {
            __Build = __Build.Trim();
            __ItemName = __ItemName.Trim();
            __Config = __Config.Trim();
            DBContext _dbContext = new DBContext();
            DataTable dt = new DataTable();
            if (checkDB)
            {
                dt = _dbContext.LoadDataTable(_DBNAME, new[] { "ItemName", "Build", "Config" },
                    new[] { __ItemName, __Build, __Config });
                if (dt.Rows.Count > 0)
                {
                    throw new DataException("Data is valid, Do you want continue?");
                }
            }
            else
            {
                dt = _dbContext.GetTableStructure(_DBNAME);
            }

            DataRow dr = dt.NewRow();
            dr["ItemName"] = __ItemName;
            dr["Build"] = __Build;
            dr["Config"] = __Config;
            dr["Data"] = getDataToJson();
            dt.Rows.Add(dr);

            _dbContext.BuckDataTable(dt, _DBNAME, new[] { "ItemName", "Build", "Config" }, null, "ID");
        }

        /// <summary>
        /// Lấy dữ liệu từ json
        /// </summary>
        /// <returns></returns>
        public string getDataToJson()
        {
            List<string> data = new List<string>();
            foreach (string key in listcategory)
            {
                data.Add($"{key}\u200B{__STATUS[key]}");
            }

            return string.Join("\u200F", data);
        }

        /// <summary>
        /// Lấy data từ json
        /// </summary>
        /// <param name="json"></param>
        /// <returns></returns>
        public void getDataFromJson(string json)
        {
            __STATUS = new Dictionary<string, string>();
            listcategory = new List<string>();
            string[] split = json.Split('\u200F');
            foreach (string item in split)
            {
                string[] splitItem = item.Split('\u200B');
                listcategory.Add(splitItem[0]);
                __STATUS.Add(splitItem[0], splitItem[1]);
            }
        }

        /// <summary>
        /// Lấy dữ liệu từ data lên datatable
        /// </summary>
        /// <returns></returns>
        public DataTable getData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("No");
            dt.Columns.Add("Name Sheet");
            dt.Columns.Add("Status Data");
            int i = 1;
            foreach (string category in listcategory)
            {
                DataRow dr = dt.NewRow();
                dr["No"] = i++;
                dr["Name Sheet"] = category;
                if (__STATUS.TryGetValue(category, out string value))
                {
                    dr["Status Data"] = value;
                }
                else
                {
                    dr["Status Data"] = "None";
                }

                dt.Rows.Add(dr);
            }

            return dt;
        }


        /// <summary>
        /// Lấy bảng dữ liệu để hiển thị
        /// </summary>
        /// <returns></returns>
        public static DataTable GetTableNameDefault()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("No");
            dt.Columns.Add("Name Sheet");
            dt.Columns.Add("Status Data");
            int i = 1;
            foreach (string category in categories)
            {
                DataRow dr = dt.NewRow();
                dr["No"] = i++;
                dr["Name Sheet"] = category;
                dr["Status Data"] = "";
                dt.Rows.Add(dr);
            }

            return dt;
        }

        public void Load(string itemName, string build, string config)
        {
            __ItemName = itemName.Trim();
            __Build = build.Trim();
            __Config = config.Trim();

            DBContext _dbContext = new DBContext();
            DataTable dt = _dbContext.LoadDataTable(_DBNAME, new[] { "ItemName", "Build", "Config" },
                new[] { itemName, build, config });
            if (dt.Rows.Count <= 0)
            {
                throw new Exception("Data not found");
            }

            string json = dt.Rows[0]["Data"].ToString();
            getDataFromJson(json);
        }

        public void checkInfo()
        {
            if (string.IsNullOrEmpty(__ItemName) || string.IsNullOrEmpty(__Build) || string.IsNullOrEmpty(__Config))
            {
                throw new Exception("ItemName, Build, config cannot be null or empty");
            }
        }

        public void CreateNew(string itemName, string build, string config)
        {
            __ItemName = itemName.Trim();
            __Build = build.Trim();
            __Config = config.Trim();
            checkInfo();

            initData();

            Save();
        }

        public void EmptyData()
        {
            __ItemName = "";
            __Build = "";
            __Config = "";
            listcategory = new List<string>();
            __STATUS = new Dictionary<string, string>();
        }

        #region Getsharekey

        // Khai báo tĩnh (static readonly) ở cấp độ class để khởi tạo 1 lần duy nhất
        private static readonly HashSet<string> _singleCheckGroup =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "flex bending", "flex bending - ict data", "thermal cycling & bending",
                "tc & bending - ict data", "heat soak & bending", "hs & bending - ict data",
                "thermal cycling", "thermal cycling - ict data", "heat soak and recovery",
                "heat soak & recovery - ict data", "thermal shock", "thermal shock - ict data"
            };

        private static readonly HashSet<string> _peelTestGroup =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "peel Test", "peel test without sus"
            };

        private static readonly HashSet<string> _singleCheckGroupIcln =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "table of contents", "assy yield", "packaging", "bar code verification"
            };

        private static readonly HashSet<string> _onProduct = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "liner peel test (on product)", "psa peel test (on product)"
        };

        public static string GetSharedKey(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
                return sheetName;

            // Chỉ gọi Trim(), bỏ hoàn toàn ToLower()
            string normalizedSheetName = sheetName.Trim();

            if (_singleCheckGroup.Contains(normalizedSheetName))
            {
                return "ORT";
            }

            if (_onProduct.Contains(normalizedSheetName))
            {
                return "ONPRODUCT";
            }

            if (_singleCheckGroupIcln.Contains(normalizedSheetName))
            {
                return "ITEMCODE_LOTNO_CHECK";
            }

            if (_singleCheckGroup.Contains(normalizedSheetName))
            {
                return "PEEL_TEST_SHEET";
            }

            return sheetName;
        }

        #endregion

        public static string getCollectFrom(string category)
        {
            string ok2Ship = "OK2SHIP";
            string ok2Build = "OK2BUILD";
            string ok2Build_Ok2Ship = "Excel OK2Build + lot ok2ship";
            string ort = "ORT";
            switch (category)
            {
                case "FAI": // check valid
                    return ok2Ship;
                case "Assy Yield": // check valid
                    return ok2Ship;
                case "Table of Contents": // check valid
                    return ok2Ship;
                case "OQC Test":
                    return ok2Ship;
                case "Cross Section":
                    return ok2Ship;
                case "GAP Connector":
                    return ok2Build;
                case "Peel Test":
                    return ok2Build;
                case "Peel Test without SUS":
                    return ok2Ship;
                case "Pull Test":
                    return ok2Build;
                case "(IQC Unmating) Pull Test":
                    return ok2Ship;
                case "Shear test":
                    return ok2Ship;
                case "OQC B2B Mating-Unmating":
                    return ok2Ship;
                case "IQC Liner peeling (Coupon)":

                    return ok2Ship;
                case "IQC PSA peeling (Coupon)":
                    return ok2Ship;
                case "Liner peel test (On product)":
                    return ok2Ship;
                case "PSA peel test (On product)":
                    return ok2Ship;
                case "Air bubble btw Liner-PSA":
                    return ok2Ship;
                case "Air bubble btw PSA-FPC":
                    return ok2Ship;
                case "ORT-Assy":
                    return ok2Ship;
                case "Flex bending":
                    return ok2Ship;
                case "Flex bending - ICT data":
                    return ok2Ship;
                case "Flex bending - X - Ray picture":
                    return ok2Ship;
                case "Thermal Cycling & bending":
                    return ok2Ship;
                case "TC & bending - ICT data ":
                    return ok2Ship;
                case "TC & bending - X - Ray picture":
                    return ok2Ship;
                case "Heat soak & bending":
                    return ok2Ship;
                case "HS & bending - ICT data":
                    return ok2Ship;
                case "HS & bending - X - Ray picture":
                    return ok2Ship;
                case "Thermal Cycling":
                    return ok2Ship;
                case "Thermal Cycling - ICT data":
                    return ok2Ship;
                case "Heat Soak and Recovery":
                    return ok2Ship;
                case "Heat soak & Recovery - ICT data":
                    return ok2Ship;
                case "Thermal Shock":
                    return ok2Ship;
                case "Thermal Shock - ICT data":
                    return ok2Ship;
                case "Environment Test":
                    return ok2Ship;
                case "Impedance":
                    return ok2Ship;
                case "ACF":
                    return ok2Ship;
                case "SEM BSE & Binarization":
                    return ok2Ship;
                case "Bar Code Verification":
                    return ok2Ship;
                case "Packaging":
                    return ok2Ship;
                case "":
                    return ok2Ship;
                default:
                    return "";
            }
        }

        public static string checkNas(string itemName, string build, string config)
        {
            List<string> list = new List<string>();
            string destinationFilePath = new ExportProcess().getExportLocation();
            destinationFilePath = Path.Combine(destinationFilePath,
                $"{itemName}_{build}_{config}");
            if (!Directory.Exists(destinationFilePath))
            {
                return "";
            }

            List<string> listfile = Directory.GetFiles(destinationFilePath).ToList();
            foreach (string item in listfile)
            {
                string name = FileFolderRepository.GetFileNameWithoutExtension(item);
                list.Add($"{name}\u200BOK");
            }

            return string.Join("\u200F", list);
        }

        public static bool checkCategory(string category, Dictionary<string, string> info, out string status,
            out string valueActtact)
        {
            valueActtact = "";
            status = "Data ERR";

            try
            {
                string itemcode = "", lotNo = "", maker = "", itemName = "", build = "", config = "";

                switch (category)
                {
                    case "NONE_GROUP":
                        status = "None";
                        return true;
                        break;
                    case "NAS_GROUP":
                        if (info.TryGetValue("ItemName", out itemName)
                            && info.TryGetValue("Build", out build)
                            && info.TryGetValue("Config", out config))
                        {
                            status = checkNas(itemName, build, config);
                            return true;
                        }

                        break;
                    case "ORT": // kiểm tra các mục ort
                        if (info.TryGetValue("ItemCode", out itemcode)
                            && info.TryGetValue("LotNo", out lotNo)
                            && info.TryGetValue("Maker", out maker))
                        {
                            status = bending_check(itemcode, lotNo, maker);
                            valueActtact = $"({itemcode}-{lotNo}-{maker})";
                            return true;
                        }

                        break;
                    case "ITEMCODE_LOTNO_CHECK": // kiểm tra trường hợp có mặt itemcode lotno
                        if (info.TryGetValue("ItemCode", out itemcode)
                            && info.TryGetValue("LotNo", out lotNo))
                        {
                            status = itemCodeLotNo_Check(itemcode, lotNo);

                            valueActtact = $"({itemcode}-{lotNo})";
                            return true;
                        }

                        break;
                    case "ONPRODUCT":
                        if (info.TryGetValue("ItemCode", out itemcode)
                            && info.TryGetValue("LotNo", out lotNo))
                        {
                            status = onproduct_check(itemcode, lotNo);

                            valueActtact = $"({itemcode}-{lotNo})";
                            return true;
                        }

                        break;

                    case "FAI":
                        if (info.TryGetValue("ItemCode", out itemcode)
                            && info.TryGetValue("LotNo", out lotNo))
                        {
                            status = FAI_check(itemcode, lotNo);

                            valueActtact = $"{itemcode}-{lotNo}";
                            return true;
                        }

                        break;
                    case "Cross Section":
                        if (info.TryGetValue("ItemCode", out itemcode)
                            && info.TryGetValue("LotNo", out lotNo))
                        {
                            status = crossSection_check(itemcode, lotNo);
                            valueActtact = $"{itemcode}-{lotNo}";
                            return true;
                        }

                        break;

                    case "OQC Test":
                        status = "No Implement";
                        return true;
                    case "GAP Connector":
                        if (info.TryGetValue("ItemCode", out itemcode)
                            && info.TryGetValue("LotNo", out lotNo))
                        {
                            status = gapconnector_check(itemcode, lotNo);
                            valueActtact = $"{itemcode}-{lotNo}";
                            return true;
                        }

                        break;

                    case "PEEL_TEST_SHEET":
                        if (info.TryGetValue("ItemCode", out itemcode)
                            && info.TryGetValue("LotNo", out lotNo))
                        {
                            status = peelTest_check(itemcode, lotNo);
                            valueActtact = $"{itemcode}-{lotNo}";
                            return true;
                        }

                        break;
                    case "SEM BSE & Binarization":
                        if (info.TryGetValue("ItemCode", out itemcode)
                            && info.TryGetValue("LotNo", out lotNo))
                        {
                            status = sem_check(itemcode, lotNo);
                            valueActtact = $"{itemcode}-{lotNo}";
                            return true;
                        }

                        break;
                    case "Environment Test":
                        if (info.TryGetValue("ItemCode", out itemcode)
                            && info.TryGetValue("LotNo", out lotNo)
                            && info.TryGetValue("Maker", out maker))
                        {
                            status = environemnt_check(itemcode, lotNo, maker);
                            valueActtact = $"{itemcode}-{lotNo}-{maker}";
                            return true;
                        }

                        break;
                    case "Pull Test":
                        if (info.TryGetValue("ItemCode", out itemcode)
                            && info.TryGetValue("LotNo", out lotNo))
                        {
                            status = pullTest_check(itemcode, lotNo);
                            valueActtact = $"{itemcode}-{lotNo}-{maker}";
                            return true;
                        }

                        break;
                    case "(IQC Unmating) Pull Test":
                        if (info.TryGetValue("ItemCode", out itemcode)
                            && info.TryGetValue("LotNo", out lotNo))
                        {
                            status = IQC_pullTest_check(itemcode, lotNo);
                            valueActtact = $"{itemcode}-{lotNo}-{maker}";
                            return true;
                        }

                        break;
                    case "Shear test":
                        if (info.TryGetValue("ItemCode", out itemcode)
                            && info.TryGetValue("LotNo", out lotNo))
                        {
                            status = shearTest_check(itemcode, lotNo);
                            valueActtact = $"{itemcode}-{lotNo}-{maker}";
                            return true;
                        }

                        break;
                    case "OQC B2B Mating-Unmating":
                        if (info.TryGetValue("ItemCode", out itemcode)
                            && info.TryGetValue("LotNo", out lotNo))
                        {
                            status = oqcummating_check(itemcode, lotNo);
                            valueActtact = $"{itemcode}-{lotNo}-{maker}";
                            return true;
                        }

                        break;
                    case "IQC Liner peeling (Coupon)":
                        if (info.TryGetValue("ItemCode", out itemcode)
                            && info.TryGetValue("LotNo", out lotNo))
                        {
                            status = iqcliner_check(itemcode, lotNo);
                            valueActtact = $"{itemcode}-{lotNo}-{maker}";
                            return true;
                        }

                        break;
                    case "IQC PSA peeling (Coupon)":
                        if (info.TryGetValue("ItemCode", out itemcode)
                            && info.TryGetValue("LotNo", out lotNo))
                        {
                            status = iqcpsa_check(itemcode, lotNo);
                            valueActtact = $"{itemcode}-{lotNo}-{maker}";
                            return true;
                        }

                        break;
                    case "Air bubble btw Liner-PSA":
                    case "Air bubble btw PSA-FPC":
                    case "ORT-Assy":
                    case "Impedance":
                    case "ACF":
                    default:
                        status = "Category No Implement function Check";
                        return true;
                }
            }
            catch
            {
            }

            return false;
        }

        #region Check Method

        private static string MSG_NODATA = "No data";
        private static string MSG_OK = "OK";

        private static string environemnt_check(string itemCode, string lotNo, string maker)
        {
            string status = "";
            DataTable dt = new DataTable();
            using (DBContext db = new DBContext())
            {
                dt = db.LoadDataTable("ENVIRONMENT_EN_DURANCE", new[] { "ItemCode", "LotNo", "Maker" },
                    new[] { itemCode, lotNo, maker });
                if (dt.Rows.Count > 0)
                {
                    status = MSG_OK;
                }
            }

            return status;
        }

        private static string sem_check(string itemCode, string lotNo)
        {
            string status = "";
            DataTable dt = new DataTable();
            using (DBContext db = new DBContext())
            {
                dt = db.LoadDataTable("SEM_BSE_Binarization_Logfile_NAS", new[] { "ItemCode", "LotNo" },
                    new[] { itemCode, lotNo });
                if (dt.Rows.Count > 0)
                {
                    status = MSG_OK;
                }
            }

            return status;
        }

        private static string pullTest_check(string itemCode, string lotNo)
        {
            string status = "";
            DataTable dt = new DataTable();
            using (DBContext db = new DBContext())
            {
                dt = db.LoadDataTable("IQC_UNMATING_PULL_TEST_NAS", new[] { "ItemCode", "LotNo", "Sheet" },
                    new[] { itemCode, lotNo, "NPI" });
                if (dt.Rows.Count > 0)
                {
                    status = MSG_OK;
                }
            }

            return status;
        }

        private static string iqcpsa_check(string itemCode, string lotNo)
        {
            string status = "";
            DataTable dt = new DataTable();
            using (DBContext db = new DBContext())
            {
                dt = db.LoadDataTable("IQC_PSA_PEELING_COUPON_NAS", new[] { "ItemCode", "LotNo", "Sheet" },
                    new[] { itemCode, lotNo, "NPI" });
                if (dt.Rows.Count > 0)
                {
                    status = MSG_OK;
                }
            }

            return status;
        }

        private static string iqcliner_check(string itemCode, string lotNo)
        {
            string status = "";
            DataTable dt = new DataTable();
            using (DBContext db = new DBContext())
            {
                dt = db.LoadDataTable("IQC_LINER_PEELING_COUPON_NAS", new[] { "ItemCode", "LotNo", "Sheet" },
                    new[] { itemCode, lotNo, "NPI" });
                if (dt.Rows.Count > 0)
                {
                    status = MSG_OK;
                }
            }

            return status;
        }

        private static string oqcummating_check(string itemCode, string lotNo)
        {
            string status = "";
            DataTable dt = new DataTable();
            using (DBContext db = new DBContext())
            {
                dt = db.LoadDataTable("OQC_B2B_Mating_Unmating_NAS", new[] { "ItemCode", "LotNo", "Sheet" },
                    new[] { itemCode, lotNo, "NPI" });
                if (dt.Rows.Count > 0)
                {
                    status = MSG_OK;
                }
            }

            return status;
        }

        private static string shearTest_check(string itemCode, string lotNo)
        {
            string status = "";
            DataTable dt = new DataTable();
            using (DBContext db = new DBContext())
            {
                dt = db.LoadDataTable("SHEAR_TEST_NAS", new[] { "ItemCode", "LotNo", "Sheet" },
                    new[] { itemCode, lotNo, "NPI" });
                if (dt.Rows.Count > 0)
                {
                    status = MSG_OK;
                }
            }

            return status;
        }

        private static string IQC_pullTest_check(string itemCode, string lotNo)
        {
            string status = "";
            DataTable dt = new DataTable();
            using (DBContext db = new DBContext())
            {
                dt = db.LoadDataTable("MATING_PULL_TEST_NAS", new[] { "ItemCode", "LotNo", "Sheet" },
                    new[] { itemCode, lotNo, "NPI" });
                if (dt.Rows.Count > 0)
                {
                    status = MSG_OK;
                }
            }


            return status;
        }

        private static string peelTest_check(string itemCode, string lotNo)
        {
            string status = "";
            string status1 = "";
            DataTable dt = new DataTable();
            using (DBContext db = new DBContext())
            {
                dt = db.LoadDataTable("PEEL_TEST_NAS", new[] { "ItemCode", "LotNo", "Sheet" },
                    new[] { itemCode, lotNo, "NPI" });
                if (dt.Rows.Count > 0)
                {
                    status = MSG_OK;
                }

                dt = db.LoadDataTable("PEEL_TEST_WITHOUT_SUS_NAS", new[] { "ItemCode", "LotNo", "Sheet" },
                    new[] { itemCode, lotNo, "NPI" });
                if (dt.Rows.Count > 0)
                {
                    status1 = MSG_OK;
                }
            }


            return $"peel Test\u200B{status}\u200Fpeel test without sus\u200B{status1}";
        }

        private static string gapconnector_check(string itemCode, string lotNo)
        {
            string status = "";
            DataTable dt = new DataTable();
            using (DBContext db = new DBContext())
            {
                dt = db.LoadDataTable("GAP_CONNECTOR_NAS", new[] { "ItemCode", "LotNo", "Sheet" },
                    new[] { itemCode, lotNo, "NPI" });
            }

            if (dt.Rows.Count > 0)
            {
                status = MSG_OK;
            }

            return status;
        }

        private static string crossSection_check(string itemCode, string lotNo)
        {
            string status = "";
            DataTable dt = new DataTable();
            using (DBContext db = new DBContext())
            {
                dt = db.LoadDataTable("CROSS_SECTION_NAS", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            }

            if (dt.Rows.Count > 0)
            {
                status = MSG_OK;
            }

            return status;
        }

        private static string FAI_check(string itemCode, string lotn)
        {
            string status = "";
            DataTable dt = new DataTable();
            using (DBContext db = new DBContext())
            {
                dt = db.LoadDataTable("FAI_Auto", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotn });
            }

            if (dt.Rows.Count > 0)
            {
                status = MSG_OK;
            }

            return status;
        }

        private static string onproduct_check(string itemCode, string lotNo)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim().PadLeft(5, '0');

            PeelTestOnProductService _service = new PeelTestOnProductService(itemCode, lotNo);
            _service._TYPE = "Displacement";
            _service.LoadData(false);

            int fullLiner = 0, fullPsa = 0;

            // Dùng HashSet để tự động loại bỏ các ID trùng lặp bị NG
            HashSet<string> ngLinerIds = new HashSet<string>();
            HashSet<string> ngPsaIds = new HashSet<string>();

            foreach (var kvp in _service._DIC)
            {
                bool isLiner = kvp.Key.Contains("LINER");

                if (isLiner)
                {
                    fullLiner = kvp.Value.Rows.Count;
                }
                else
                {
                    fullPsa = kvp.Value.Rows.Count;
                }

                foreach (DataRow row in kvp.Value.Rows)
                {
                    // So sánh trực tiếp chuỗi, không nối chuỗi (string interpolation)
                    if (row["JudgementForce"]?.ToString() == "NG")
                    {
                        string id = row["ID"]?.ToString();
                        if (!string.IsNullOrEmpty(id))
                        {
                            if (isLiner)
                                ngLinerIds.Add(id);
                            else
                                ngPsaIds.Add(id);
                        }
                    }
                }
            }

            return
                $"psa peel test (on product)\u200B{ngPsaIds.Count}/{fullPsa}\u200Fliner peel test (on product)\u200B{ngLinerIds.Count}/{fullLiner}";
        }

        private static string itemCodeLotNo_Check(string itemCode, string lotNo)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim().PadLeft(5, '0');
            DataTable dt = new DataTable();
            using (DBContext _db = new DBContext())
            {
                string sqlQuery = $"" +
                                  $"WITH FoundRecords AS (" +
                                  $"SELECT 'Table of Contents' AS ExistsCheck FROM TABLE_OF_CONTENT WHERE ItemCode = @ItemCode AND LotNo = @LotNo " +
                                  $"UNION ALL " +
                                  $"SELECT 'Assy Yield' AS ExistsCheck FROM ASSY_YIELD WHERE ItemCode = @ItemCode AND LotNo = @LotNo " +
                                  $"UNION ALL " +
                                  $"SELECT 'Bar Code Verification' AS ExistsCheck FROM BAR_CODE_VERIFICATION WHERE ItemCode = @ItemCode AND LotNo = @LotNo " +
                                  $"UNION ALL " +
                                  $"SELECT 'Packaging' AS ExistsCheck FROM PACKAGING_LOGFILE WHERE ItemCode = @ItemCode" +
                                  $"),AllTables AS (" +
                                  $" SELECT 'Shear test' AS TableName " +
                                  $" UNION ALL SELECT 'Table of Contents'  " +
                                  $" UNION ALL SELECT 'Assy Yield' " +
                                  $" UNION ALL SELECT 'Bar Code Verification' " +
                                  $" UNION ALL SELECT 'Packaging' " +
                                  $")" +
                                  $" SELECT A.TableName AS NameTable, CASE WHEN F.ExistsCheck IS NOT NULL THEN '1' ELSE '0' END AS Status" +
                                  $" FROM AllTables AS A LEFT JOIN FoundRecords AS F ON A.TableName = F.ExistsCheck";

                SqlParameter[] sqlParameters = new SqlParameter[]
                {
                    new SqlParameter("@ItemCode", itemCode.PadRight(10)),
                    new SqlParameter("@LotNo", lotNo.PadRight(10))
                };
                dt = _db.ExecuteQueryToDataTable(sqlQuery, sqlParameters);
            }

            int acf = 0;
            List<string> res = new List<string>();
            foreach (DataRow row in dt.Rows)
            {
                string nameTable = row["NameTable"].ToString();
                string status = row["Status"].ToString();
                bool prime = status.Contains("1");
                if (prime == true)
                {
                    res.Add($"{nameTable}\u200BOK");
                }
                else
                {
                    res.Add($"{nameTable}\u200BNG");
                }
            }

            return string.Join("\u200F", res);
        }

        private static string bending_check(string itemCode, string lotNo, string maker)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim().PadLeft(5, '0');
            maker = maker.Trim();
            string status = "";
            DataTable dt = new DataTable();
            using (DBContext db = new DBContext())
            {
                dt = db.LoadDataTable("BENDING_TEST_DATA", new[] { "ItemCode", "LotNo", "Maker" },
                    new[] { itemCode, lotNo, maker });
            }

            if (dt != null && dt.Rows.Count > 0)
            {
                string checking = dt.Rows[0]["Result"].ToString();
                if (string.IsNullOrEmpty(checking))
                {
                    return MSG_NODATA;
                }

                return checking;
            }

            return MSG_NODATA;
        }

        #endregion
    }
}