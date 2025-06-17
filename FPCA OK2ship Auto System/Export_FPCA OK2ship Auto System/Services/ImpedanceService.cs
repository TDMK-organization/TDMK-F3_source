using Export_FPCA_OK2ship_Auto_System.Repositories;
using OfficeOpenXml.Drawing;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Authentication;
using System.Text;
using System.Threading.Tasks;

namespace Export_FPCA_OK2ship_Auto_System.Services
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


        public void Export(ExcelWorksheet worksheet, string itemCode, string lotNo)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
            {
                throw new Exception("Hãy nhập itemcode và lotno!");
            }
            Dictionary<string, DataTable> dic = Load(itemCode, lotNo);

            ExportProcess exportProcess = new ExportProcess();

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
                    string address = addressZ.Split('-')[0];
                    int i = 0;
                    string region = dataTableZ.Rows[0]["Region"].ToString();
                    foreach (DataRow item in dataTableZ.Rows)
                    {
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


        }
    }
}


