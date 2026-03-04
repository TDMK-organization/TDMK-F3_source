using Export_FPCA_OK2ship_Auto_System.Repositories;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Export_FPCA_OK2ship_Auto_System.Services
{

    public class ExportService
    {
        private DBContext _db = new DBContext();
        private string[] categories =
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
            "OQC Test",
            "Cross section",
            "GAP Connector",
            "Peel Test",
            "(Mating) Pull Test",
            "IQC Liner peeling (Coupon)",
            "IQC PSA peeling (Coupon)",
            "Shear test",
            "Liner peel test (On product)",
            "Air bubble btw Liner-PSA",
            "PSA peel test (On product)",
            "Air bubble btw PSA-FPC",
            "(IQC Unmating) Pull Test",
            "OQC B2B Mating_Unmating",
            "ORT-Assy",
            "Flex bending",
            "Thermal Cycling & bending",
            "Heat Soak & bending",
            "X_Ray picture",
            "Heat Soak and Recovery",
            "Thermal Cycling",
            "Thermal Shock",
            "Environment en_durance",
            "Impedance",
            "Switch Quality",
            "ACF",
            "SEM BSE & Binarization",
            "Bar Code Verification",
            "Packaging",
            "Mishandling test",
            "Process flow",
            "Process Comparison",
        };
        public string[] getCategories()
        {
            return categories;
        }
        public Dictionary<string, int> _dicCategory = new Dictionary<string, int>();
        public DataTable setupDGV()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Category", typeof(string));
            dt.Columns.Add("Status Manual", typeof(string));
            dt.Columns.Add("Export manual", typeof(bool));
            dt.Columns.Add("StatusDB", typeof(string));
            dt.Columns.Add("Export form DB", typeof(bool));
            dt.Columns.Add("Export status", typeof(string));
            foreach (string item in categories)
            {
                _dicCategory.Add(item, dt.Rows.Count + 1);
                dt.Rows.Add(dt.Rows.Count + 1, item, "None", false, "None", false, "");
            }
            return dt;

        }
        public bool? checkingDataCategory(string category, string itemCode, string lotNo)
        {
            try
            {

                switch (category)
                {
                    case "GAP Connector":
                        _db = new DBContext();
                        using (DataTable dt = _db.LoadDataTable("GAP_CONNECTOR_NAS", new[] { "ItemCode", "LotNo", "Sheet" }, new[] { itemCode.PadRight(10), lotNo.PadRight(10), "NPI" }, new[] { "ItemCode", "LotNo", "Sheet" }))
                        {
                            if (dt.Rows.Count > 0)
                            {

                                return true;
                            }

                        }
                        break;

                    case "Shear test":
                        _db = new DBContext();
                        using (DataTable dt = _db.LoadDataTable("SHEAR_TEST_NAS", new[] { "ItemCode", "LotNo", "Sheet" }, new[] { itemCode.PadRight(10), lotNo.PadRight(10), "NPI" }, new[] { "ItemCode", "LotNo", "Sheet" }))
                        {
                            if (dt.Rows.Count > 0)
                            {

                                return true;
                            }

                        }
                        break;
                    case "ACF":
                        _db = new DBContext();
                        using (DataTable dt = _db.LoadDataTable("ACF_BONDING_NAS", new[] { "ItemCode", "LotNo", "Sheet" }, new[] { itemCode.PadRight(10), lotNo.PadRight(10), "NPI" }, new[] { "ItemCode", "LotNo", "Sheet" }))
                        {
                            if (dt.Rows.Count > 0)
                            {

                                return true;
                            }

                        }
                        break;
                    case "Peel Test":
                        _db = new DBContext();
                        using (DataTable dt = _db.LoadDataTable("PEEL_TEST_NAS", new[] { "ItemCode", "LotNo", "Sheet" }, new[] { itemCode.PadRight(10), lotNo.PadRight(10), "NPI" }, new[] { "ItemCode", "LotNo", "Sheet" }))
                        {
                            if (dt.Rows.Count > 0)
                            {

                                return true;
                            }

                        }
                        break;
                    case "(Mating) Pull Test":
                        _db = new DBContext();
                        using (DataTable dt = _db.LoadDataTable("MATING_PULL_TEST_NAS", new[] { "ItemCode", "LotNo", "Sheet" }, new[] { itemCode.PadRight(10), lotNo.PadRight(10), "NPI" }, new[] { "ItemCode", "LotNo", "Sheet" }))
                        {
                            if (dt.Rows.Count > 0)
                            {

                                return true;
                            }

                        }
                        break;
                    default:
                        return null;

                }
            }
            catch
            {
                return null;
            }
            return false;
        }

        public List<string> checkingDatabase(string itemCode, string lotNo)
        {
            List<string> list = new List<string>();
            #region OLD
            //foreach (string item in categories)
            //{
            //    bool? prime = checkingDataCategory(item, itemCode, lotNo);
            //    switch (prime)
            //    {
            //        case true:
            //            list.Add($"{item}-StatusDB-OK");
            //            break;
            //        case false:
            //            list.Add($"{item}-StatusDB-No Data");
            //            break;
            //        case null:
            //            list.Add($"{item}-StatusDB-No Valid");
            //            break;
            //    }
            //}
            #endregion
            #region NEW

            string sqlQuery = $"" +
                $"WITH FoundRecords AS (" +
                $"SELECT 'Table of Contents' AS ExistsCheck FROM TABLE_OF_CONTENT WHERE ItemCode = @ItemCode AND LotNo = @LotNo " +
                $"UNION ALL " +
                $"SELECT 'Assy Yield' AS ExistsCheck FROM ASSY_YIELD WHERE ItemCode = @ItemCode AND LotNo = @LotNo " +
                $"UNION ALL " +
                $"SELECT 'Cross section' AS ExistsCheck FROM CROSS_SECTION_NAS WHERE ItemCode = @ItemCode AND LotNo = @LotNo " +
                $"UNION ALL " +
                $"SELECT 'GAP Connector' AS ExistsCheck FROM GAP_CONNECTOR_NAS WHERE ItemCode = @ItemCode AND LotNo = @LotNo AND Sheet = 'NPI'" +
                $"UNION ALL " +
                $"SELECT 'Peel Test' AS ExistsCheck FROM PEEL_TEST_NAS WHERE ItemCode = @ItemCode AND LotNo = @LotNo AND Sheet = 'NPI' " +
                $"UNION ALL " +
                $"SELECT '(Mating) Pull Test' AS ExistsCheck FROM MATING_PULL_TEST_NAS WHERE ItemCode = @ItemCode AND LotNo = @LotNo AND Sheet = 'NPI' " +
                $"UNION ALL " +
                $"SELECT 'Shear test' AS ExistsCheck FROM SHEAR_TEST_NAS WHERE ItemCode = @ItemCode AND LotNo = @LotNo AND Sheet = 'NPI' " +
                $"UNION ALL " +
                $"SELECT 'Liner\\PSA peel test (On product)' AS ExistsCheck FROM PT_ONPRODUCT_NAS WHERE ItemCode = @ItemCode AND LotNo = @LotNo " +
                $"UNION ALL " +
                $"SELECT 'OQC B2B Mating_Unmating' AS ExistsCheck FROM OQC_B2B_Mating_Unmating_NAS WHERE ItemCode = @ItemCode AND LotNo = @LotNo " +
                $"UNION ALL " +
                $"SELECT 'X-Ray picture' AS ExistsCheck FROM XRAY WHERE ItemCode = @ItemCode AND LotNo = @LotNo " +
                $"UNION ALL " +
                $"SELECT 'Heat Soak and Recovery' AS ExistsCheck FROM TC_HS_TS WHERE ItemCode = @ItemCode AND LotNo = @LotNo AND [Type] = 'HS' " +
                $"UNION ALL " +
                $"SELECT 'Thermal Cycling' AS ExistsCheck FROM TC_HS_TS WHERE ItemCode = @ItemCode AND LotNo = @LotNo AND [Type] = 'TC'  " +
                $"UNION ALL " +
                $"SELECT 'Thermal Shock' AS ExistsCheck FROM TC_HS_TS WHERE ItemCode = @ItemCode AND LotNo = @LotNo AND [Type] = 'TS' " +
                $"UNION ALL " +
                $"SELECT 'Environment en_durance' AS ExistsCheck FROM ENVIRONMENT_EN_DURANCE WHERE ItemCode = @ItemCode AND LotNo = @LotNo " +
                $"UNION ALL " +
                $"SELECT 'Impedance' AS ExistsCheck FROM IMPEDANCE_VAL_LOGFILE WHERE ItemCode = @ItemCode AND LotNo = @LotNo " +
                $"UNION ALL " +
                $"SELECT 'ACF_BONDING_NAS' AS ExistsCheck FROM ACF_BONDING_NAS WHERE ItemCode = @ItemCode AND LotNo = @LotNo " +
                $"UNION ALL " +
                $"SELECT 'ACF_FLATNESS_NAS' AS ExistsCheck FROM ACF_BONDING_NAS WHERE ItemCode = @ItemCode AND LotNo = @LotNo " +
                $"UNION ALL " +
                $"SELECT 'Roughness_NAS' AS ExistsCheck FROM Roughness_NAS WHERE ItemCode = @ItemCode AND LotNo = @LotNo " +
                $"UNION ALL " +
                $"SELECT 'ACF_WCA' AS ExistsCheck FROM ACF_WCA WHERE ItemCode = @ItemCode AND LotNo = @LotNo and [Type] = 'NPI' " +
                $"UNION ALL " +
                $"SELECT 'SEM BSE & Binarization' AS ExistsCheck FROM SEM_BSE_Binarization_Logfile_NAS WHERE ItemCode = @ItemCode AND LotNo = @LotNo " +
                $"UNION ALL " +
                $"SELECT 'Bar Code Verification' AS ExistsCheck FROM BAR_CODE_VERIFICATION WHERE ItemCode = @ItemCode AND LotNo = @LotNo " +
                $"UNION ALL " +
                $"SELECT 'Packaging' AS ExistsCheck FROM PACKAGING_LOGFILE WHERE ItemCode = @ItemCode" +
                $"),AllTables AS (" +
                $" SELECT 'Shear test' AS TableName " +
                $" UNION ALL SELECT 'Table of Contents'  " +
                $" UNION ALL SELECT 'Assy Yield' " +
                $" UNION ALL SELECT 'Cross section' " +
                $" UNION ALL SELECT 'GAP Connector' " +
                $" UNION ALL SELECT '(Mating) Pull Test' " +
                $" UNION ALL SELECT 'Peel Test' " +
                $" UNION ALL SELECT 'Liner\\PSA peel test (On product)' " +
                $" UNION ALL SELECT 'OQC B2B Mating_Unmating' " +
                $" UNION ALL SELECT 'ACF_BONDING_NAS' " +
                $" UNION ALL SELECT 'ACF_FLATNESS_NAS' " +
                $" UNION ALL SELECT 'Roughness_NAS' " +
                $" UNION ALL SELECT 'ACF_WCA' " +
                $" UNION ALL SELECT 'Thermal Shock'" +
                $" UNION ALL SELECT 'Thermal Cycling' " +
                $" UNION ALL SELECT 'Heat Soak and Recovery' " +
                $" UNION ALL SELECT 'X-Ray picture' " +
                $" UNION ALL SELECT 'Environment en_durance'" +
                $" UNION ALL SELECT 'Impedance' " +
                $" UNION ALL SELECT 'Bar Code Verification' " +
                $" UNION ALL SELECT 'Packaging' " +
                $" UNION ALL SELECT 'SEM BSE & Binarization'" +
                $")" +
                $" SELECT A.TableName AS NameTable, CASE WHEN F.ExistsCheck IS NOT NULL THEN '1' ELSE '0' END AS Status" +
                $" FROM AllTables AS A LEFT JOIN FoundRecords AS F ON A.TableName = F.ExistsCheck";
            SqlParameter[] sqlParameters = new SqlParameter[]
            {
                new SqlParameter("@ItemCode", itemCode.PadRight(10)),
                new SqlParameter("@LotNo", lotNo.PadRight(10))
            };
            DataTable dt = _db.ExecuteQueryToDataTable(sqlQuery, sqlParameters);
            int acf = 0;
            foreach (DataRow row in dt.Rows)
            {
                string nameTable = row["NameTable"].ToString();
                string status = row["Status"].ToString();
                bool prime = status.Contains("1");
                //fillnametable
                List<string> nameTables = new List<string>();
                switch (nameTable)
                {
                    case "ACF_FLATNESS_NAS":
                    case "Roughness_NAS":
                    case "ACF_BONDING_NAS":
                        acf += prime == true ? 1 : 0;
                        break;
                    case "ACF_WCA":
                        acf += prime == true ? 1 : 0;
                        list.Add($"ACF-StatusDB-OK {acf}/4");

                        break;
                    case "X-Ray picture":
                        nameTables.Add($"X_Ray picture");
                        break;
                    case "Liner\\PSA peel test (On product)":
                        nameTables.Add($"Liner peel test (On product)");
                        nameTables.Add($"PSA peel test (On product)");
                        break;
                    default:
                        nameTables.Add(nameTable);
                        break;
                }
                foreach (string item in nameTables)
                {
                    switch (prime)
                    {
                        case true:
                            list.Add($"{item}-StatusDB-OK");
                            break;
                        case false:
                            list.Add($"{item}-StatusDB-No Data");
                            break;
                        default:
                            list.Add($"{item}-StatusDB-no valid");
                            break;
                    }
                }
            }

            #endregion
            return list;
        }
        private string _EXPORT_KEY = "ExportLE";
        public string ExportOneByOne(string itemCode, string lotNo, string category)
        {
            string msg = "";
            switch (category)
            {
                case "OQC B2B Mating_Unmating":
                    OQCB2BMatingUnmatting OQCServcie = new OQCB2BMatingUnmatting();
                    msg = $"{category}-{_EXPORT_KEY}-{OQCServcie.Export(itemCode, lotNo)}";
                    break;
                case "ACF":
                    ACFService acfServcies = new ACFService();
                    msg = $"{category}-{_EXPORT_KEY}-{acfServcies.Export(itemCode, lotNo)}";
                    break;
                case "Packaging":
                    PackagingService packService = new PackagingService();
                    msg = $"{category}-{_EXPORT_KEY}-{packService.Export(itemCode, lotNo)}";
                    break;

                case "Bar Code Verification":
                    BarCodeVertification barService = new BarCodeVertification();
                    msg = $"{category}-{_EXPORT_KEY}-{barService.Export(itemCode, lotNo)}";
                    break;
                case "Environment en_durance":
                    EEDService eService = new EEDService(null);
                    msg = $"{category}-{_EXPORT_KEY}-{eService.Export(itemCode, lotNo)}";
                    break;

                case "SEM BSE & Binarization":
                    SEMServices semService = new SEMServices();
                    try
                    {
                        msg = $"{category}-{_EXPORT_KEY}-{semService.Export(itemCode, lotNo, true)}";
                    }
                    catch (Exception ex)
                    {
                        msg = $"{category}-{_EXPORT_KEY}-{ex.Message}";
                    }
                    break;
                case "Assy Yield":
                    AssyYieldService ayService = new AssyYieldService(null);
                    msg = $"{category}-{_EXPORT_KEY}-{ayService.Export(itemCode, lotNo)}";
                    break;
                case "Cross section":
                    CrossSectionService crossService = new CrossSectionService();
                    msg = $"{category}-{_EXPORT_KEY}-{crossService.Export(itemCode, lotNo)}";
                    break;
                case "Table of Contents":
                    TableOfContentService tocService = new TableOfContentService();
                    msg = $"{category}-{_EXPORT_KEY}-{tocService.Export(itemCode, lotNo)}";
                    break;
                case "GAP Connector":
                    GAPConnectorService service = new GAPConnectorService();
                    msg = $"{category}-{_EXPORT_KEY}-{service.Export(itemCode, lotNo)}";
                    break;
                case "Peel Test":
                case "(Mating) Pull Test":
                case "Shear test":
                    PeelPullShearService peelService = new PeelPullShearService();
                    msg = $"{category}-{_EXPORT_KEY}-{peelService.Export(itemCode, lotNo, category)}";
                    break;
                case "Heat Soak and Recovery":
                case "Thermal Cycling":
                case "Thermal Shock":
                    TCHSTSService TCHSTSService = new TCHSTSService();
                    msg = $"{category}-{_EXPORT_KEY}-{TCHSTSService.Export(itemCode, lotNo, category)}";
                    break;
                default:
                    msg = $"{category}-{_EXPORT_KEY}-No Implemented";
                    break;
            }
            return msg;
        }
        public List<string> ExportOneByOne(string itemCode, string lotNo, string[] listCategoryExportFormDB)
        {
            List<string> msgList = new List<string>();
            // Export lẻ từng sheet đã chọn
            foreach (string item in listCategoryExportFormDB)
            {
                msgList.Add(ExportOneByOne(itemCode, lotNo, item));
            }



            return msgList;
        }
        private string _CHECK_MANUAL = "CheckManual";
        public string CheckManualOneByOne(string itemCode, string lotNo, string category, string locationCheck)
        {
            string msg = "";
            try
            {
                bool prime = FileFolderRepository.checkLocationIsValid($"{locationCheck}\\{category}");
                if (prime)
                {
                    string nameFile = $"{locationCheck}\\{category}\\{itemCode}_{lotNo}";
                    if (File.Exists($"{nameFile}.xlsm"))
                    {
                        msg = "OK";
                    }
                    if (File.Exists($"{nameFile}.xlsx"))
                    {
                        msg = "OK";
                    }
                    return msg;
                }
                else
                {
                    msg = "NO DATA";
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return msg;
        }
        public List<string> CheckManual(string itemCode, string lotNo)
        {
            ExportProcess process = new ExportProcess();
            string locationCheck = process.getExportLocation();
            List<string> msgList = new List<string>();
            foreach (string item in categories)
            {
                msgList.Add($"{item}-{_CHECK_MANUAL}-{CheckManualOneByOne(itemCode, lotNo, item, locationCheck)}");
            }

            return msgList;
        }
        /// <summary>
        /// Tổng hợp báo cáo
        /// </summary>
        /// <param name="itemCode"></param>
        /// <param name="lotNo"></param>
        /// <param name="listExportDB"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public List<string> ExportAll(string itemCode, string lotNo, string[] listExport)
        {
            List<string> msgList = new List<string>();

            //Tim format tong
            using (ExportProcess process = new ExportProcess())
            {
                using (ExcelPackage package = process.FindFormatWithItemCode(itemCode))
                {
                    // Quet tung sheet nếu có trong danh sách thì thay thế sheet vào
                    foreach (string category in listExport)
                    {
                        ExcelWorksheet oldWorksheet = package.Workbook.Worksheets[category];
                        if (oldWorksheet != null)
                        {
                            package.Workbook.Worksheets.Delete(oldWorksheet);
                        }
                    }
                    foreach (string category in listExport)
                    {
                        string msg = "";
                        bool prime = FileFolderRepository.checkLocationIsValid($"{process.getExportLocation()}\\{category}");
                        if (prime)
                        {
                            try
                            {
                                string nameFile = $"{process.getExportLocation()}\\{category}\\{itemCode}_{lotNo}";
                                if (File.Exists($"{nameFile}.xlsm") || File.Exists($"{nameFile}.xlsx"))
                                {
                                    nameFile = nameFile + (File.Exists($"{nameFile}.xlsm") ? ".xlsm" : ".xlsx");
                                    ExcelPackage.LicenseContext = LicenseContext.Commercial;

                                    using (ExcelPackage packageLe = new ExcelPackage(nameFile))
                                    {
                                        ExcelPackage.LicenseContext = LicenseContext.Commercial;

                                        List<ExcelWorksheet> workSheets = new List<ExcelWorksheet>();
                                        foreach (ExcelWorksheet worksheet in packageLe.Workbook.Worksheets)
                                        {

                                            ExcelWorksheet sheetDestination = package.Workbook.Worksheets.Add($"{worksheet.Name}", worksheet);
                                            sheetDestination.TabColor = Color.Green;
                                        }


                                        msg = "OK";
                                    }

                                }
                                else
                                {
                                    msg = "File Không tồn tại";
                                }
                            }
                            catch (Exception ex)
                            {
                                msg = ex.Message;
                            }
                        }
                        //Kiem tra va tra ket qua
                        msgList.Add($"{category}-{_EXPORT_KEY}-{msg}");
                    }
                    // sắp xếp các sheet về vị trí cũ
                    SortWorkSheet(package);
                    // Lưu lại sheet vào đường dẫn
                    process.SaveExcelPackage(package, $"Export all {itemCode}_{lotNo}");
                }
            }


            return msgList;
        }
        string[] listCategoriesZ = new[]
        {
            "Coverpage","Rev History","User Guideline","Low CPK Action","Declaration","Table of Contents","Deviation Summary","Assy Yield","FAI-1","FAI-2","FAI-3","SPC-1","SPC-2","SPC-3","FAI without parentheses","FAI with parentheses","OQC Test","Cross Section","GAP Connector","Peel Test","(Mating) Pull Test","Shear test","Liner peel test (On product)","Air bubble btw Liner-PSA","PSA peel test (On product)","Air bubble btw PSA-FPC","(IQC Unmating) Pull Test","OQC B2B Mating-Unmating","ORT-Assy ","Flex bending","Thermal Cycling & bending","Heat soak & bending","X-Ray picture","Heat Soak and Recovery","Thermal Cycling","Thermal Shock","Environment en-durance","Impedance","Switch Quality","ACF","SEM BSE & Binarization ","Bar Code Verification","Packaging","Mishandling test","Process flow","Process Comparison"
        };
        public void SortWorkSheet(ExcelPackage package)
        {

            List<string> categorys = listCategoriesZ.ToList();
            List<string> nameSheet = package.Workbook.Worksheets.AsEnumerable().Select(x => x.Name).ToList();
            int i = 0;
            foreach (string ws in categorys)
            {
                int index = nameSheet.IndexOf(ws);
                if (index != -1)
                {
                    package.Workbook.Worksheets.MoveToEnd(ws);
                }
            }
        }

        public string SetupManual(string itemCode, string lotNo, string location, string category, bool takeAll)
        {
            string msg = "";
            if (string.IsNullOrEmpty(location))
            {
                throw new Exception($"Location không được để trống!");
            }
            using (ExportProcess process = new ExportProcess())
            {

                string locationRoot = process.getExportLocation() + $"\\{category}";
                if (!Directory.Exists(locationRoot))
                {
                    // 2. Nếu không tồn tại, thì tạo
                    try
                    {
                        Directory.CreateDirectory(locationRoot);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
                string nameFile = $"{locationRoot}\\{itemCode}_{lotNo}";
                if (File.Exists($"{nameFile}.xlsm"))
                {
                    msg = "New Change";
                    nameFile = nameFile + ".xlsm";
                }
                if (File.Exists($"{nameFile}.xlsx"))
                {
                    msg = "New Change";
                    nameFile = nameFile + ".xlsx";
                }
                if (msg == "New Change")
                {
                    File.Delete(nameFile);
                }

                using (ExcelPackage package = process.OpenFileExcel(location))
                {
                    List<string> nameSheet = new List<string>();
                    if (!takeAll)
                    {
                        switch (category)
                        {
                            case "FAI":
                                string[] str = new[] { "FAI-1", "FAI-2", "SPC-1", "SPC-2", "Dimension without parenthes", "Dimension with parenthes" };
                                int i = 0;
                                foreach (string item in str)
                                {
                                    if (package.Workbook.Worksheets.Count() >= i)
                                    {
                                        break;
                                    }
                                    ExcelWorksheet workSheet1 = package.Workbook.Worksheets[i++];
                                    workSheet1.Name = item;
                                    nameSheet.Add(item);
                                }
                                break;
                            case "Peel test":
                                string[] strZ = new[] { "Peel Test", "Peel Test Without SUS" };
                                int iZ = 0;
                                foreach (string item in strZ)
                                {
                                    if (package.Workbook.Worksheets.Count() >= iZ)
                                    {
                                        break;
                                    }
                                    ExcelWorksheet workSheet1 = package.Workbook.Worksheets[iZ++];
                                    workSheet1.Name = item;
                                    nameSheet.Add(item);
                                }
                                break;
                            default:
                                ExcelWorksheet workSheet = package.Workbook.Worksheets[0];
                                workSheet.Name = category;
                                nameSheet.Add(category);

                                break;
                        }
                    }
                    else
                    {
                        foreach (ExcelWorksheet item in package.Workbook.Worksheets)
                        {
                            string nameWS = item.Name;
                            nameSheet.Add(nameWS);

                        }
                    }
                    process.SaveExcelWorksheet(package, $"{string.Join(":", nameSheet)}", $"{itemCode}_{lotNo}", "NPI", false, ExportProcess.GetFileExtension(location));

                }
            }
            return msg;
        }
    }
}
