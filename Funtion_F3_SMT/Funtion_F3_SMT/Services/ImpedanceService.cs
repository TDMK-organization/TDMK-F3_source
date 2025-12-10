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
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();

            if (string.IsNullOrEmpty(lotNo) || string.IsNullOrEmpty(itemCode))
            {
                throw new ArgumentException("ItemCode and LotNo cannot be empty.");
            }
            Dictionary<string, DataTable> dic = new Dictionary<string, DataTable>();



            //string filter_str = TDMK_Code.filter_str(new string[] { "ItemCode", "LotNo" }, new string[] { txtItemCode.Text, txtLotNo.Text });
            //DGV_Impedance_Summary.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_VAL", filter_str);

            dic.Add("IMPEDANCE_VAL", _dbContext.LoadDataTable("IMPEDANCE_VAL", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo }));


            //DGV_Graph.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", filter_str);
            dic.Add("IMPEDANCE_GRAPH", _dbContext.LoadDataTable("IMPEDANCE_GRAPH", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo }));

            //DGV_Impedance_Spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
            dic.Add("IMPEDANCE_SPEC", _dbContext.LoadDataTable("IMPEDANCE_SPEC", new[] { "ItemCode" }, new[] { itemCode }));

            //DGV_Tracewidth_spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
            dic.Add("TRACEWIDTH_SPEC", _dbContext.LoadDataTable("TRACEWIDTH_SPEC", new[] { "ItemCode" }, new[] { itemCode }));



            //DGV_VHX_Data.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_VAL", filter_str);
            dic.Add("TRACEWIDTH_VAL", _dbContext.LoadDataTable("TRACEWIDTH_VAL", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo }));

            //DGV_Image_Tracewidth.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_IMAGE", filter_str);
            dic.Add("TRACEWIDTH_IMAGE", _dbContext.LoadDataTable("TRACEWIDTH_IMAGE", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo }));


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

            dic.Add("IMPEDANCE_VAL", _dbContext.LoadDataTable("IMPEDANCE_VAL", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo }));
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
            dic.Add("IMPEDANCE_VAL", ChangeItemCodeLotNo(dataTable, icNew, lnNew));

            dataTable = _db.LoadDataTable("IMPEDANCE_GRAPH", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            //DGV_Graph.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_GRAPH", filter_str);
            dic.Add("IMPEDANCE_GRAPH", ChangeItemCodeLotNo(dataTable, icNew, lnNew));
            dataTable = _db.LoadDataTable("IMPEDANCE_SPEC", new[] { "ItemCode" }, new[] { itemCode });
            //DGV_Impedance_Spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "IMPEDANCE_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
            dic.Add("IMPEDANCE_SPEC", ChangeItemCodeLotNo(dataTable, icNew));

            //DGV_Tracewidth_spec.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_SPEC", TDMK_Code.filter_str(new string[] { "ItemCode" }, new string[] { txtItemCode.Text }));
            dataTable = _db.LoadDataTable("TRACEWIDTH_SPEC", new[] { "ItemCode" }, new[] { itemCode });
            dic.Add("TRACEWIDTH_SPEC", ChangeItemCodeLotNo(dataTable, icNew));


            dataTable = _db.LoadDataTable("TRACEWIDTH_VAL", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            //DGV_VHX_Data.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_VAL", filter_str);
            dic.Add("TRACEWIDTH_VAL", ChangeItemCodeLotNo(dataTable, icNew, lnNew));

            dataTable = _db.LoadDataTable("TRACEWIDTH_IMAGE", new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            //DGV_Image_Tracewidth.DataSource = TDMK_Code.Datatable_Filter(sqlcon, "TRACEWIDTH_IMAGE", filter_str);
            dic.Add("TRACEWIDTH_IMAGE", ChangeItemCodeLotNo(dataTable, icNew, lnNew));


            return dic;
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
                    res += _dbContext.BuckDataTable(dataTable, item, list.ToArray(), null, "ID");
                }
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
                    #endregion

                    #region Data Impedance 1
                    if (dic.TryGetValue("IMPEDANCE_VAL", out DataTable dataTableZ))
                    {

                        if (_dic.TryGetValue("Actual Impedance", out string addressZ))
                        {
                            int i = 0, j = -1;
                            string region = dataTableZ.Rows[0]["Region"].ToString();
                            string address = "";
                            foreach (DataRow item in dataTableZ.Rows)
                            {
                                if (item["Pcs_No"].ToString().Trim().Equals("1"))
                                {
                                    j++;
                                    i = 0;
                                    address = addressZ.Split('-')[j];
                                }
                                if (!item["Region"].ToString().Equals(region))
                                {
                                    region = item["Region"].ToString();
                                    i = 0;
                                    address = addressZ.Split('-')[int.Parse(region) - 1];
                                }
                                i++;
                                if (double.TryParse(item["Data"].ToString(), out double data))
                                {
                                    string add = ExportProcess.AddRow(address, i);
                                    data = Math.Round(data, 2);
                                    worksheet.Cells[ExportProcess.AddRow(address, i)].Value = data;
                                }

                            }
                        }
                    }

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
