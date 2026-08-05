using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using OfficeOpenXml;
using OK2SHIP_SMT.Repositories;

namespace OK2SHIP_SMT.Services
{
    public class ERS_NetSpecService
    {
        public DataTable _DATA = new DataTable();

        public ERS_NetSpecService()
        {
            _DATA = getDataTableStructor();
            _DATA.Rows.Add(" ");
            
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
            return dt;
        }

        public void getData(string location, string itemCode, string maker)
        {
            _DATA = getDataTableStructor();
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
                        new[] { "Pin1", "Pin2", "Net name", "LowLimit", "HighLimit", "Min DCR", "Max DCR" });
                    int i = 1;
                    int colPin1 = worksheet.Cells[dic["Pin1"]].Start.Column;
                    int colPin2 = worksheet.Cells[dic["Pin2"]].Start.Column;
                    int colNetname = worksheet.Cells[dic["Net name"]].Start.Column;
                    int colLowLimit = worksheet.Cells[dic["LowLimit"]].Start.Column;
                    int colHighLimit = worksheet.Cells[dic["HighLimit"]].Start.Column;
                    int colMinDCR = worksheet.Cells[dic["Min DCR"]].Start.Column;
                    int colMaxDCR = worksheet.Cells[dic["Max DCR"]].Start.Column;
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

                        DataRow rowValue = _DATA.NewRow();
                        rowValue["Net Name"] = valueNetName;
                        rowValue["Pin1"] = worksheet.Cells[row + i, colPin1].Value.ToString();
                        rowValue["Pin2"] = worksheet.Cells[row + i, colPin2].Value.ToString();
                        rowValue["Low Limit"] = worksheet.Cells[row + i, colLowLimit].Value.ToString();
                        rowValue["High Limit"] = worksheet.Cells[row + i, colHighLimit].Value.ToString();
                        rowValue["Min DCR"] = worksheet.Cells[row + i, colMinDCR].Value.ToString();
                        rowValue["Max DCR"] = worksheet.Cells[row + i, colMaxDCR].Value.ToString();
                        _DATA.Rows.Add(rowValue);
                        i++;
                    }
                }
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
            dr["Data"] = ConverterService.DataTableToJson(_DATA);
            dt.Rows.Add(dr);
            
            _DBCONTEXT.BuckDataTable(dt, _NAMETABLE, new[] { "ItemCode", "Maker" }, null, "ID");
        }

        public void LoadData(string itemCode, string maker)
        {
            
            DataTable dt = _DBCONTEXT.LoadDataTable(_NAMETABLE,  new[] { "ItemCode", "Maker" }, new[] { itemCode, maker });
            if (dt.Rows.Count <= 0)
            {
                _DATA = getDataTableStructor();
                _DATA.Rows.Add(" ");
                throw new Exception("Không có dữ liệu!");
            }

            DataRow row = dt.Rows[0];

            string data = row["Data"].ToString();
            _DATA = ConverterService.JsonToDataTable(data);
             
        }
    }
}