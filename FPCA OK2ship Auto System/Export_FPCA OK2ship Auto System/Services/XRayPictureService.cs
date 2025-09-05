using Export_FPCA_OK2ship_Auto_System.Repositories;
using Export_FPCA_OK2ship_Auto_System.Services.TDMK_services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Export_FPCA_OK2ship_Auto_System.Services
{
    public class XRayPictureService
    {
        private readonly string _NAME_SQL = "XRAY";
        private DBContext _dbContext = new DBContext();
        public string ConvertDataTable(DataTable datatable, DataTable dataTableImage, ref int id, int area)
        {
            DataTable resDT = new DataTable();
            //Add Column
            foreach (DataColumn column in datatable.Columns)
            {
                if (column.DataType.FullName == "System.Drawing.Image")
                {
                    resDT.Columns.Add(column.ColumnName + "&CONVERTER", typeof(string));
                }
                else if (column.DataType.FullName == "System.Byte[]")
                {
                    resDT.Columns.Add(column.ColumnName + "&CONVERTER", typeof(string));
                }
                else
                {
                    resDT.Columns.Add(column.ColumnName, column.DataType);
                }
            }
            //add row
            foreach (DataRow row in datatable.Rows)
            {
                DataRow rowres = resDT.NewRow();
                foreach (DataColumn col in resDT.Columns)
                {
                    if (col.ColumnName.Contains("&CONVERTER"))
                    {
                        DataRow newRow = dataTableImage.NewRow();
                        string colName = col.ColumnName.Replace("&CONVERTER", "");
                        var image = row[colName];
                        if (datatable.Columns[colName].DataType.FullName != "System.Byte[]")
                        {
                            image = TDMK_ImageConverter.ImageToByteArray((Image)row[colName], ImageFormat.Jpeg);

                        }

                        rowres[col.ColumnName] = id;
                        newRow["ID"] = id++;
                        newRow["Image"] = image;
                        newRow["Area"] = area;
                        dataTableImage.Rows.Add(newRow);
                    }
                    else
                    {
                        rowres[col.ColumnName] = row[col.ColumnName];
                    }
                }
                resDT.Rows.Add(rowres);
            }

            return TDMK_ConverterService.DataTableToJson(resDT);
        }
        public DataTable ConvertDataTable(DataTable datatable, DataTable dataTableImage)
        {
            DataTable res = new DataTable();
            foreach (DataColumn col in datatable.Columns)
            {
                if (col.ColumnName.Contains("&CONVERTER"))
                {
                    res.Columns.Add(col.ColumnName.Replace("&CONVERTER", ""), typeof(byte[]));
                }
                else
                {
                    res.Columns.Add(col.ColumnName);
                }
            }
            foreach (DataRow row in datatable.Rows)
            {
                DataRow rowZ = res.NewRow();
                foreach (DataColumn col in datatable.Columns)
                {

                    if (col.ColumnName.Contains("&CONVERTER"))
                    {
                        string name = col.ColumnName.Replace("&CONVERTER", "");
                        int point = int.Parse(row[col].ToString()) - int.Parse(dataTableImage.Rows[0]["ID"].ToString());
                        byte[] img = (byte[])dataTableImage.Rows[point]["Image"];
                        rowZ[name] = img;
                    }
                    else
                    {
                        rowZ[col.ColumnName] = row[col.ColumnName];
                    }
                }
                res.Rows.Add(rowZ);
            }

            return res;
        }
        public DataTable Load(string itemCode, string lotNo, string type)
        {
            if (string.IsNullOrEmpty(type.Trim()))
            {
                throw new Exception("Hãy chọn type!");
            }
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            if (string.IsNullOrEmpty(itemCode) && string.IsNullOrEmpty(lotNo))
            {
                throw new Exception("Không được để trống itemcode lotno");
            }
            DataTable dataTable = _dbContext.LoadDataTable(_NAME_SQL, new[] { "ItemCode", "LotNo", "Type" }, new[] { itemCode, lotNo, type }, new[] { "Area", "Data" });
            if (dataTable.Rows.Count == 0)
            {
                throw new Exception("Không tìm thấy dữ liệu");
            }
            DataTable dataImage = _dbContext.LoadDataTable(_NAME_SQL + "_IMAGE", new[] { "Area" }, new[] { dataTable.Rows[0]["Area"].ToString() }, new[] { "Image", "ID" });
            DataTable dataTableRes = ConvertDataTable(TDMK_ConverterService.JsonToDataTable((string)dataTable.Rows[0]["Data"]), dataImage);

            return dataTableRes;

        }
        public void Export(ExcelWorksheet workSheet, string itemCode, string lotNo, string type)
        {

            ExportProcess exportProcess = new ExportProcess();


            DataTable dataTable = Load(itemCode, lotNo, type);
            string[] header = new[] { "#", "Bending", "Flex SN" };
            IDictionary<string, string> keyValuePairs = ExportProcess.FindAddressByText(workSheet, header);
            //
            Dictionary<string, string> headler = new Dictionary<string, string>();

            //create dic for bending
            foreach (string item in keyValuePairs["Bending"].Split('-'))
            {
                string key = workSheet.Cells[item].Value.ToString().Replace("Bending", "").Replace(" ", "");
                string[] addz = keyValuePairs["#"].Split('-');
                List<string> list = new List<string>();
                foreach (string item1 in addz)
                {
                    list.Add(workSheet.Cells[workSheet.Cells[item].End.Row, workSheet.Cells[item1].End.Column].Address);
                }
                headler.Add(key, string.Join("-", list));
            }

            foreach (DataRow row in dataTable.Rows)
            {
                string area = row["Area"].ToString().Replace(" ", "").Replace("B", "");
                if (headler.TryGetValue(area, out string value))
                {
                    if (value != "")
                    {
                        var address = workSheet.Cells[value.Split('-')[0]];
                        ExportProcess.InsertImageToCell(workSheet, address, (byte[])row["Image"], $"{Guid.NewGuid()}");
                        headler[area] = headler[area].Replace($"{address.Address}-", "");
                        workSheet.Cells[ExportProcess.AddRow(address.Address, 1)].Value = row["Result"].ToString();
                    }
                }
            }
        }

    }
}


