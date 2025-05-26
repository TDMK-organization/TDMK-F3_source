using OfficeOpenXml;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
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
    public class XRayPictureService
    {
        private readonly string _NAME_SQL = "XRAY";
        private DBContext _dbContext = new DBContext();
        public bool CheckNameFile(string location, string itemCode, string lotNo)
        {
            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
            {
                throw new Exception("Không được để itemcode lotno trống");
            }
            if (location != null)
            {
                if (!FileFolderRepository.checkLocationIsValid(location))
                {
                    throw new Exception("Địa chỉ không tồn tại");
                }
                string[] lo = FileFolderRepository.GetFolderName(location).Split(new[] { '-', '_' });
                string _itemCode = lo[1];
                string _lotNo = lo[2] + (lo[3].Count() <= 2 ? lo[3] : "");
                if (!(itemCode.Equals(_itemCode) && _lotNo.Equals(lotNo)))
                {
                    return false;
                }
            }
            return true;
        }
        public int Save(DataTable dataTable, string itemCode, string lotNo, int prime = -1)
        {
            if (dataTable.Rows.Count <= 0)
            {
                throw new Exception("Hãy get data!");
            }
            int res = 0;
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();

            int area = 0;
            if (prime == -1)
            {
                DataTable log = _dbContext.LoadDataTable(_NAME_SQL, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo }, new[] { "ID", "Area" });
                if (log.Rows.Count > 0)
                {
                    string z = log.Rows[0]["Area"].ToString();
                    throw new Exception($" 1234 - {int.Parse(z)} - Đã tồn tại dữ liệu bạn có muốn tiếp tục!");
                }
            }

            area = _dbContext.GetID(_NAME_SQL + "_Image") + 1;
            DataTable resDataTable = _dbContext.GetTableStructure(_NAME_SQL);
            DataTable dataTableImage = _dbContext.GetTableStructure(_NAME_SQL + "_Image");
            DataRow dr = resDataTable.NewRow();
            int id = area;
            string json = ConvertDataTable(dataTable, dataTableImage, ref id, area);
            dr["ItemCode"] = itemCode;
            dr["LotNo"] = lotNo;
            dr["Data"] = json;
            dr["Area"] = area;
            resDataTable.Rows.Add(dr);
            res += _dbContext.BuckDataTable(dataTableImage, _NAME_SQL + "_Image", new[] { "Area" });
            _dbContext.DeleteData(_NAME_SQL + "_Image", "Area", new[] { prime.ToString() });
            res += _dbContext.BuckDataTable(resDataTable, _NAME_SQL, new[] { "ItemCode", "LotNo" }, null, "ID");
            return res;
        }
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

            return ConverterService.DataTableToJson(resDT);
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
        public DataTable Read(string location, string itemCode, string lotNo)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            if (!CheckNameFile(location, itemCode, lotNo))
            {
                throw new Exception("Vấn đề itemcode lotno");
            }
            string[] subFolder = FileFolderRepository.GetSubFolders(location);
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("ID");
            dataTable.Columns.Add("Area");
            dataTable.Columns.Add("Result");
            dataTable.Columns.Add("Image", typeof(byte[]));
            foreach (string folder in subFolder)
            {
                IList<KeyValuePair<Image, string>> list = FileFolderRepository.ListAllPictureInAFolder(folder, ".png");
                string[] folderName = FileFolderRepository.GetFolderName(folder).Split('-');
                string ar = folderName[folderName.Length - 1];
                int id = 1;
                foreach (var item in list)
                {
                    DataRow row = dataTable.NewRow();
                    row["Area"] = ar;
                    row["ID"] = id++;
                    row["Image"] = TDMK_ImageConverter.ImageToByteArray(item.Key, ImageFormat.Png);
                    dataTable.Rows.Add(row);
                }
            }
            return dataTable;
        }

        public DataTable Load(string itemCode, string lotNo)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            if (string.IsNullOrEmpty(itemCode) && string.IsNullOrEmpty(lotNo))
            {
                throw new Exception("Không được để trống itemcode lotno");
            }
            DataTable dataTable = _dbContext.LoadDataTable(_NAME_SQL, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo }, new[] { "Area", "Data" });
            if (dataTable.Rows.Count == 0)
            {
                throw new Exception("Không tìm thấy dữ liệu");
            }
            DataTable dataImage = _dbContext.LoadDataTable(_NAME_SQL + "_IMAGE", new[] { "Area" }, new[] { dataTable.Rows[0]["Area"].ToString() }, new[] { "Image", "ID" });
            DataTable dataTableRes = ConvertDataTable(ConverterService.JsonToDataTable((string)dataTable.Rows[0]["Data"]), dataImage);

            return dataTableRes;

        }

        public string Export(string itemCode, string lotNo)
        {
            ExportProcess exportProcess = new ExportProcess();
            using (ExcelPackage ex = exportProcess.FindFormatProcess("X-Ray picture", itemCode, lotNo))
            {
                using (ExcelWorksheet workSheet = exportProcess.FindSheet(ex, "X-Ray picture"))
                {
                    DataTable dataTable = Load(itemCode, lotNo);
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
                    exportProcess.SaveExcelWorksheet(ex, "X-Ray picture", $"{itemCode.Trim()}-{lotNo.Trim()}");
                }
            }
            return "Export thành công!";
        }

    }
}
