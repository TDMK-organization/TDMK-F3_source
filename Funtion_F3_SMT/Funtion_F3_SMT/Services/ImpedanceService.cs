using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Security.Authentication;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZedGraph;

namespace OK2SHIP_SMT.Services
{
    public class ImpedanceService
    {

        private DBContext _dbContext = new DBContext();
        public Dictionary<string, DataTable> Load(string itemCode, string lotNo)
        {
            string msg = "";
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();

            if (string.IsNullOrEmpty(lotNo) || string.IsNullOrEmpty(itemCode))
            {
                throw new ArgumentException("ItemCode and LotNo cannot be empty.");
            }
            Dictionary<string, DataTable> dic = new Dictionary<string, DataTable>();

            DataTable dt = _dbContext.LoadDataTable("IMPEDANCE_NAS", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            DataTable dt_SPEC = _dbContext.LoadDataTable("IMPEDANCE_SPEC_NAS", new[] { "ItemCode" }, new[] { itemCode });
            if (dt.Rows.Count < 0)
            {
                throw new Exception("Không tồn tại dữ liệu!");
            }

            if (dt_SPEC.Rows.Count < 0)
            {
                msg += "Dữ liệu SPEC không tồn tại \n";
            }

            try
            {
                DataTable impedance_val = ConverterService.JsonToDataTable(dt.Rows[0]["IMPEDANCE_VAL"].ToString());
                dic.Add("IMPEDANCE_VAL", impedance_val);
            }
            catch (Exception ex)
            {
                msg += "Impedance_Value error: " + ex.Message + "\n";
            }

            try
            {
                DataTable impedance_g = ConverterService.JsonToDataTable(dt.Rows[0]["IMPEDANCE_GRAPH"].ToString());
                NasRepository nas = new NasRepository();
                nas.MergeDataTable(impedance_g, "IMPEDANCE_GRAPH", itemCode, lotNo, dt.Rows[0]["impedance_IMG"].ToString());
                //DGV_Graph.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", filter_str);
                dic.Add("IMPEDANCE_GRAPH", impedance_g);
            }
            catch (Exception ex)
            {
                msg += "Impedance graph error: " + ex.Message + "\n";
            }

            try
            {
                DataTable impedance_val = ConverterService.JsonToDataTable(dt_SPEC.Rows[0]["IMPEDANCE_SPEC"].ToString());
                dic.Add("IMPEDANCE_SPEC", impedance_val);
            }
            catch (Exception ex)
            {
                msg += "impedance spec error: " + ex.Message + "\n";
            }

            try
            {
                DataTable impedance_val = ConverterService.JsonToDataTable(dt_SPEC.Rows[0]["TRACEWIDTH_SPEC"].ToString());
                dic.Add("TRACEWIDTH_SPEC", impedance_val);
            }
            catch (Exception ex)
            {
                msg += "tracewidth spec error: " + ex.Message + "\n";
            }


            try
            {
                DataTable impedance_val = ConverterService.JsonToDataTable(dt.Rows[0]["TRACEWIDTH_VAL"].ToString());
                dic.Add("TRACEWIDTH_VAL", impedance_val);
            }
            catch (Exception ex)
            {
                msg += "tracewidth Value error: " + ex.Message + "\n";
            }
            try
            {
                DataTable impedance_g = ConverterService.JsonToDataTable(dt.Rows[0]["TRACEWIDTH_IMAGE"].ToString());
                NasRepository nas = new NasRepository();
                nas.MergeDataTable(impedance_g, "TRACEWIDTH_IMAGE", itemCode, lotNo, dt.Rows[0]["tracewidth_IMG"].ToString());
                dic.Add("TRACEWIDTH_IMAGE", impedance_g);
            }
            catch (Exception ex)
            {
                msg += "Tracewidth image error: " + ex.Message + "\n";
            }

            return dic;
        }
        public bool CheckDataIsNotNull(string itemCode, string lotNo)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();

            if (string.IsNullOrEmpty(lotNo) || string.IsNullOrEmpty(itemCode))
            {
                throw new ArgumentException("ItemCode and LotNo cannot be empty.");
            }

            Dictionary<string, DataTable> dic = new Dictionary<string, DataTable>();



            //string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            //DGV_Impedance_Summary.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_VAL", filter_str);

            dic.Add("IMPEDANCE_VAL", OverWriteDatatable(SortDataValueTable(_dbContext.LoadDataTable("IMPEDANCE_VAL", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo }))));
            if (dic["IMPEDANCE_VAL"].Rows.Count > 0)
            {
                return false;
            }

            //DGV_Graph.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", filter_str);
            dic.Add("IMPEDANCE_GRAPH", _dbContext.LoadDataTable("IMPEDANCE_GRAPH", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo }));
            if (dic["IMPEDANCE_GRAPH"].Rows.Count > 0)
            {
                return false;
            }
            //DGV_Impedance_Spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
            dic.Add("IMPEDANCE_SPEC", _dbContext.LoadDataTable("IMPEDANCE_SPEC", new[] { "ItemCode" }, new[] { itemCode }));
            if (dic["IMPEDANCE_SPEC"].Rows.Count > 0)
            {
                return false;
            }
            //DGV_Tracewidth_spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
            dic.Add("TRACEWIDTH_SPEC", _dbContext.LoadDataTable("TRACEWIDTH_SPEC", new[] { "ItemCode" }, new[] { itemCode }));
            if (dic["TRACEWIDTH_SPEC"].Rows.Count > 0)
            {
                return false;
            }


            //DGV_VHX_Data.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_VAL", filter_str);
            dic.Add("TRACEWIDTH_VAL", _dbContext.LoadDataTable("TRACEWIDTH_VAL", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo }));
            if (dic["TRACEWIDTH_VAL"].Rows.Count > 0)
            {
                return false;
            }
            //DGV_Image_Tracewidth.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_IMAGE", filter_str);
            dic.Add("TRACEWIDTH_IMAGE", _dbContext.LoadDataTable("TRACEWIDTH_IMAGE", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo }));

            if (dic["TRACEWIDTH_IMAGE"].Rows.Count > 0)
            {
                return false;
            }
            return true;
        }
        public static Dictionary<string, DataTable> LoadOldImpedanceData(string itemCode, string lotNo, string icNew, string lnNew)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();

            if (string.IsNullOrEmpty(lotNo) || string.IsNullOrEmpty(itemCode))
            {
                throw new ArgumentException("ItemCode and LotNo cannot be empty.");
            }
            //inti
            DBContext _db = new DBContext("OK2SHIP_Period2");
            Dictionary<string, DataTable> dic = new Dictionary<string, DataTable>();



            //string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            //DGV_Impedance_Summary.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_VAL", filter_str);
            DataTable dataTable = _db.LoadDataTable("IMPEDANCE_VAL", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });

            dic.Add("IMPEDANCE_VAL", OverWriteDatatable(SortDataValueTable(ChangeItemCodeLotNo(dataTable, icNew, lnNew))));

            dataTable = _db.LoadDataTable("IMPEDANCE_GRAPH", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            //DGV_Graph.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", filter_str);
            dic.Add("IMPEDANCE_GRAPH", OverWriteDatatable(ChangeItemCodeLotNo(dataTable, icNew, lnNew)));
            dataTable = _db.LoadDataTable("IMPEDANCE_SPEC", new[] { "ItemCode" }, new[] { itemCode });
            //DGV_Impedance_Spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
            dic.Add("IMPEDANCE_SPEC", OverWriteDatatable(ChangeItemCodeLotNo(dataTable, icNew)));

            //DGV_Tracewidth_spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
            dataTable = _db.LoadDataTable("TRACEWIDTH_SPEC", new[] { "ItemCode" }, new[] { itemCode });
            dic.Add("TRACEWIDTH_SPEC", OverWriteDatatable(ChangeItemCodeLotNo(dataTable, icNew)));


            dataTable = _db.LoadDataTable("TRACEWIDTH_VAL", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            //DGV_VHX_Data.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_VAL", filter_str);
            dic.Add("TRACEWIDTH_VAL", OverWriteDatatable(SortDataValueTable(ChangeItemCodeLotNo(dataTable, icNew, lnNew), true)));

            dataTable = _db.LoadDataTable("TRACEWIDTH_IMAGE", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            //DGV_Image_Tracewidth.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_IMAGE", filter_str);
            dic.Add("TRACEWIDTH_IMAGE", OverWriteDatatable(ChangeItemCodeLotNo(dataTable, icNew, lnNew)));


            return dic;
        }
        private static DataTable OverWriteDatatable(DataTable dt)
        {
            if (dt.Columns.Contains("ID"))
            {
                // 2. Chạy vòng lặp qua từng dòng của DataTable
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    // Gán giá trị bằng (i + 1) để bắt đầu từ số 1
                    dt.Rows[i]["ID"] = i + 1;
                }
            }
            return dt;
        }
        private static DataTable SortDataValueTable(DataTable dataTable, bool tracewidth = false)
        {
            if (tracewidth)
            {
                return dataTable.AsEnumerable()
               .OrderBy(row =>
               {
                   int region;
                   int.TryParse(row["Pcs_No"]?.ToString().Replace("Sample", "").Trim(), out region);
                   return region;
               })
               .ThenBy(row =>
               {
                   int pcsNo;
                   int.TryParse(row["Region"]?.ToString(), out pcsNo);
                   return pcsNo; // Nếu lỗi/rỗng sẽ trả về 0
               })
               .CopyToDataTable();
            }

            return dataTable.AsEnumerable()
                .OrderBy(row =>
                {
                    int pcsNo;
                    int.TryParse(row["Region"]?.ToString(), out pcsNo);
                    return pcsNo; // Nếu lỗi/rỗng sẽ trả về 0
                })
                .ThenBy(row =>
                {
                    int region;
                    int.TryParse(row["Pcs_No"]?.ToString(), out region);
                    return region;
                })
                .CopyToDataTable();
        }
        private static DataTable ChangeItemCodeLotNo(DataTable dataTable, string itemCode, string lotNo = null)
        {
            foreach (DataRow row in dataTable.Rows)
            {
                row["ItemCode"] = itemCode;
                if (lotNo != null)
                {
                    row["LotNo"] = lotNo;
                }
            }
            return dataTable;
        }
        public DataTable initNAStable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("ItemCode", typeof(string));
            dt.Columns.Add("LotNo", typeof(string));
            dt.Columns.Add("IMPEDANCE_GRAPH", typeof(string));
            dt.Columns.Add("IMPEDANCE_VAL", typeof(string));
            dt.Columns.Add("TRACEWIDTH_VAL", typeof(string));
            dt.Columns.Add("TRACEWIDTH_IMAGE", typeof(string));
            dt.Columns.Add("impedance_IMG", typeof(string));
            dt.Columns.Add("tracewidth_IMG", typeof(string));

            return dt;
        }
        public DataTable initNAS_SPECtable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("ItemCode", typeof(string));
            dt.Columns.Add("IMPEDANCE_SPEC", typeof(string));
            dt.Columns.Add("TRACEWIDTH_SPEC", typeof(string));
            return dt;
        }
        public int Save(string itemCode, string lotNo, Dictionary<string, DataTable> dic, bool prime = false)
        {


            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();

            if (string.IsNullOrEmpty(UserSession.Instance.User_ID))
            {
                throw new AuthenticationException("Bạn chưa đăng nhập, hãy đăng nhập!");

            }
            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
            {
                throw new Exception("Không được trống itemcode lotno");
            }


            if (prime == false)
            {
                if (CheckDataIsNotNull(itemCode, lotNo) == false)
                {
                    throw new Exception("1234 - Có dữ liệu trong DB bạn muốn ghi đè?");
                }

            }

            DataTable dt = initNAStable();
            DataTable dt_SPEC = initNAS_SPECtable();
            DataRow row = dt.NewRow();
            DataRow row_SPEC = dt_SPEC.NewRow();

            row["ItemCode"] = itemCode;
            row["LotNo"] = lotNo;
            row_SPEC["ItemCode"] = itemCode;


            int res = 0;
            foreach (string item in dic.Keys)
            {
                if (dic[item].Rows.Count == 0)
                {

                }
                else
                {
                    DataTable dataTable = dic[item];
                    List<string> list = new List<string>();
                    list.Add("ItemCode");
                    switch (item)
                    {
                        case "IMPEDANCE_GRAPH":
                        case "IMPEDANCE_VAL":
                            list.Add("LotNo");
                            break;
                        case "IMPEDANCE_SPEC":
                        case "TRACEWIDTH_SPEC":
                            if (dataTable.Columns.IndexOf("Format_Type") != -1)
                            {
                                dataTable.Columns.Remove("Format_Type");
                            }
                            break;
                        default:
                            break;
                    }
                    //convert json
                    switch (item)
                    {
                        case "IMPEDANCE_GRAPH":
                        case "TRACEWIDTH_IMAGE":
                            NasRepository nas = new NasRepository();
                            string location = nas.HandleImageDataTable(dataTable, item, itemCode, lotNo);

                            string key = "";
                            if (item.Equals("IMPEDANCE_GRAPH"))
                            {
                                key = "impedance_IMG";
                            }
                            if (item.Equals("TRACEWIDTH_IMAGE"))
                            {
                                key = "tracewidth_IMG";
                            }
                            row[key] = location;
                            row[item] = ConverterService.DataTableToJson(dataTable);
                            break;
                        case "IMPEDANCE_VAL":
                        case "TRACEWIDTH_VAL":
                            row[item] = ConverterService.DataTableToJson(dataTable);
                            break;
                        case "IMPEDANCE_SPEC":
                        case "TRACEWIDTH_SPEC":
                            row_SPEC[item] = ConverterService.DataTableToJson(dataTable);
                            break;
                        default:
                            Debugger.Break();
                            break;
                    }
                }
            }

            dt.Rows.Add(row);
            dt_SPEC.Rows.Add(row_SPEC);

            try
            {
                res += _dbContext.BuckDataTable(dt, "IMPEDANCE_NAS", new[] { "ItemCode", "LotNo" }, null, "ID");
            }
            catch (Exception ex)
            {
                throw new Exception("Error Save Impedance:" + ex.Message);
            }
            try
            {
                res += _dbContext.BuckDataTable(dt_SPEC, "IMPEDANCE_SPEC_NAS", new[] { "ItemCode" }, null, "ID");
            }
            catch (Exception ex)
            {
                throw new Exception("Error Save Impedance SPEC:" + ex.Message);
            }
            return res;
        }

        public void Export(string itemCode, string lotNo)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
            {
                throw new Exception("Hãy nhập itemcode và lotno!");
            }
            Dictionary<string, DataTable> dic = Load(itemCode, lotNo);

            ExportProcess exportProcess = new ExportProcess();
            using (ExcelPackage ex = exportProcess.FindFormatProcess("Impedance", itemCode, lotNo))
            {
                using (ExcelWorksheet worksheet = exportProcess.FindSheet(ex, "Impedance"))
                {
                    string[] header = new[] { "Actual Impedance", "Picture", "Impedance -" };
                    IDictionary<string, string> _dic = ExportProcess.FindAddressByText(worksheet, header, true);

                    #region Image

                    string[] headerZ = new[] { "Sample 1", "Sample 2", "Sample 3", "Actual Trace width" };
                    IDictionary<string, string> _dicZ = ExportProcess.FindAddressByText(worksheet, headerZ, false);
                    try
                    {

                        if (dic.TryGetValue("IMPEDANCE_GRAPH", out DataTable dataTable))
                        {
                            int ha = 1;
                            foreach (DataRow item in dataTable.Rows)
                            {
                                string samplez = $"Sample {ha++}";
                                byte[] image = (byte[])item["Image_Graph"];
                                if (_dicZ.TryGetValue(samplez, out string address))
                                {
                                    if (address.Split('-').Count() >= 2)
                                    {

                                        address = ExportProcess.AddRow(address.Split('-')[1], 2);

                                        ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[ExportProcess.getRangeBaseAddressByCellAddress(worksheet, address)], image, $"{Guid.NewGuid()}");
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception exzx)
                    {
                        throw new Exception("IMPEDANCE_GRAPH:" + exzx.Message);
                    }
                    try
                    {

                        if (dic.TryGetValue("TRACEWIDTH_IMAGE", out DataTable dataTableImage))
                        {
                            int ha = 1;
                            foreach (DataRow item in dataTableImage.Rows)
                            {
                                string samplez = $"Sample {ha++}";
                                byte[] image = (byte[])item["Image_Tracewidth"];
                                if (_dicZ.TryGetValue(samplez, out string address))
                                {
                                    if (address.Split('-').Count() >= 2)
                                    {
                                        address = ExportProcess.AddRow(address.Split('-')[0], 1);
                                        var z = worksheet.Cells[ExportProcess.getRangeBaseAddressByCellAddress(worksheet, address)];
                                        ExportProcess.InsertImageToCell(worksheet, z, image, $"{Guid.NewGuid()}");
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception exzx)
                    {
                        throw new Exception("TRACEWIDTH_IMAGE:" + exzx.Message);
                    }
                    #endregion

                    #region Data Impedance 1
                    try
                    {
                        if (dic.TryGetValue("IMPEDANCE_VAL", out DataTable dataTableZ))
                        {
                            if (_dic.TryGetValue("Actual Impedance", out string addressZ))
                            {
                                foreach (DataRow item in dataTableZ.Rows)
                                {
                                    if (int.TryParse(item["Pcs_No"].ToString().Trim(), out int pcsNo))
                                    {
                                        if (double.TryParse(item["Data"].ToString(), out double data))
                                        {
                                            if (int.TryParse(item["Region"].ToString().Trim(), out int region))
                                            {
                                                string address = addressZ.Split('-')[region - 1];
                                                data = Math.Round(data, 2);
                                                worksheet.Cells[ExportProcess.AddRow(address, pcsNo)].Value = data;
                                            }
                                            else
                                            {
                                                throw new Exception("Dữ liệu region có lỗi");
                                            }
                                        }
                                        else { throw new Exception("Dữ liệu trong cột Data có lỗi, không thể chuyển sang số thực"); }
                                    }

                                }
                            }
                        }

                    }
                    catch (Exception exZ)
                    {
                        throw new Exception("Export IMPEDANCE_VAL:" + exZ.Message);
                    }
                    try
                    {

                        if (dic.TryGetValue("TRACEWIDTH_VAL", out DataTable dataTableZ1))
                        {
                            if (_dicZ.TryGetValue("Actual Trace width", out string address))
                            {
                                address = address.Split('-')[0];
                                int i = 0;
                                foreach (DataRow item in dataTableZ1.Rows)
                                {
                                    i++;
                                    if (double.TryParse(item["Data"].ToString(), out double data))
                                    {
                                        string add = ExportProcess.AddRow(address, i);
                                        worksheet.Cells[ExportProcess.AddRow(address, i)].Value = Math.Round(data, 2);
                                    }
                                }
                            }
                            else
                            {

                            }
                        }
                    }
                    catch (Exception exZ)
                    {
                        throw new Exception("Export TRACEWIDTH_VAL:" + exZ.Message);
                    }
                    #endregion


                    exportProcess.SaveExcelWorksheet(ex, "Impedance", $"{itemCode} - {lotNo} IMPEDANCE");
                }
            }
        }

        public Dictionary<string, DataTable> GetData(string itemCode, string lotNo, string location)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            location = location.Trim();
            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo) || string.IsNullOrEmpty(location))
            {

                throw new Exception("Không được để trường nào trống");
            }
            ExportProcess exportProcess = new ExportProcess();
            Dictionary<string, DataTable> _dic = new Dictionary<string, DataTable>();
            string[] tableName = new[] { "TRACEWIDTH_IMAGE", "TRACEWIDTH_SPEC", "TRACEWIDTH_VAL", "IMPEDANCE_GRAPH", "IMPEDANCE_IMAGE", "IMPEDANCE_SPEC", "IMPEDANCE_VAL" };
            foreach (string item in tableName)
            {
                _dic.Add(item, _dbContext.GetTableStructure(item));
            }
            using (ExcelPackage package = new ExcelPackage(location))
            {
                if (package == null)
                {
                    throw new Exception("File excel have error");
                }
                using (ExcelWorksheet worksheet = package.Workbook.Worksheets["Impedance"])
                {
                    if (worksheet == null)
                    {
                        throw new Exception("Worksheet lỗi");
                    }
                    string[] healder = new[] { "Sample 1", "Sample 2", "Sample 3", "Actual Impedance", "Sample No.", "Actual Trace width (A)" };
                    IDictionary<string, string> addressDIC = ExportProcess.FindAddressByText(worksheet, healder);
                    Dictionary<string, byte[]> imageList = new Dictionary<string, byte[]>();
                    foreach (var picture in worksheet.Drawings)
                    {
                        if (picture is ExcelPicture)
                        {
                            byte[] image = ((ExcelPicture)picture).Image.ImageBytes;
                            int column = picture.From.Column + 1;
                            int row = picture.From.Row;
                            string key = worksheet.Cells[row, column].Address;
                            imageList.Add(key, image);
                        }
                    }
                    for (int i = 1; i < 4; i++)
                    {
                        if (addressDIC.TryGetValue($"Sample {i}", out string address))
                        {
                            string[] az = address.Split('-');
                            if (imageList.TryGetValue(az[1], out byte[] value))
                            {

                                DataRow row = ((DataTable)_dic["IMPEDANCE_GRAPH"]).NewRow();
                                row["ID"] = i;
                                row["ItemCode"] = itemCode;
                                row["LotNo"] = lotNo;
                                row["Pcs_No"] = i;
                                row["Region"] = 1;
                                row["Image_Graph"] = value;
                                row["Operator"] = UserSession.Instance.User_ID;
                                row["Zone"] = "Patern";
                                row["Remark"] = "GET FORM CHECKSHEET";
                                _dic["IMPEDANCE_GRAPH"].Rows.Add(row);
                            }
                            else
                            {

                            }
                            if (imageList.TryGetValue(az[0], out value))
                            {
                                DataRow row = ((DataTable)_dic["TRACEWIDTH_IMAGE"]).NewRow();
                                row["ID"] = i;
                                row["ItemCode"] = itemCode;
                                row["LotNo"] = lotNo;
                                row["Pcs_No"] = i;
                                row["Data_For"] = "A";
                                row["Image_Tracewidth"] = value;
                                row["Operator"] = UserSession.Instance.User_ID;
                                row["Time_Update"] = "Patern";
                                row["Remark"] = "GET FORM CHECKSHEET";
                                _dic["TRACEWIDTH_IMAGE"].Rows.Add(row);
                            }
                        }
                    }
                    bool prime = false;
                    if (addressDIC.TryGetValue("Actual Impedance", out string add))
                    {
                        string[] zHealder = new[] { "IMPEDANCE_VAL" };
                        int j = 0;
                        foreach (string address in add.Split('-'))
                        {
                            string addressZZ = address;

                            int i = 1;
                            while (true)
                            {
                                addressZZ = ExportProcess.AddRow(addressZZ, 1);
                                if (worksheet.Cells[addressZZ] != null && !string.IsNullOrEmpty(worksheet.Cells[addressZZ].Text))
                                {
                                    if (double.TryParse(worksheet.Cells[addressZZ].Text, out double data))
                                    {
                                        //string sampeAdd = worksheet.Cells[worksheet.Cells[addressZZ].End.Row, worksheet.Cells[addressDIC["Sample No."].Split('-')[j]].End.Column].Address;
                                        //if (string.IsNullOrEmpty(worksheet.Cells[sampeAdd].Text))
                                        //{
                                        //    break;
                                        //}
                                        DataRow row = ((DataTable)_dic[zHealder[0]]).NewRow();
                                        row["ID"] = i;
                                        row["ItemCode"] = itemCode;
                                        row["LotNo"] = lotNo;
                                        row["Pcs_No"] = i++;
                                        row["Region"] = j + (j > 0 ? 0 : 1);
                                        row["Data"] = data;
                                        row["Zone"] = "Patern";
                                        row["Remark"] = "GET FORM EXCEL";
                                        _dic["IMPEDANCE_VAL"].Rows.Add(row);
                                    }
                                    else
                                    {
                                        data = 0;
                                    }

                                }
                                else
                                {
                                    break;
                                }
                            }

                            j++;
                            if (j == 1)
                            {
                                j++;
                            }
                        }
                    }
                }
            }
            return _dic;
        }
    }
}
