using Export_FPCA_OK2ship_Auto_System.Repositories;
using Export_FPCA_OK2ship_Auto_System.Services.TDMK_services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Export_FPCA_OK2ship_Auto_System.Services
{
    public class EEDService
    {
        private DBContext _dBContext;
        private const string FORMAT_NAME = "Environment en-durance";
        private const string _TABLE_NAME = "ENVIRONMENT_EN_DURANCE";
        public EEDService(SqlConnection sqlcon)
        {
            if (sqlcon == null)
            {
                _dBContext = new DBContext();
            }
            else
            {
                _dBContext = new DBContext(sqlcon);
            }
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
        public Dictionary<string, Dictionary<string, DataTable>> Load(string itemCode, string lotNo, ref DataTable before, bool legacy = false)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
            {
                throw new Exception("ItemCode or LotNo is null");
            }
            DataTable table = _dBContext.LoadDataTable(_TABLE_NAME, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            if (table.Rows.Count <= 0)
            {
                throw new Exception("Không có dữ liệu của itemcode lotno");
            }
            string[] json = table.Rows[0]["Data"].ToString().Trim().Split('#');
            string area = json[0];
            DataTable tableImage = new DataTable();
            if (legacy)
            {
                tableImage = _dBContext.LoadDataTable(_TABLE_NAME + "_IMG", new[] { "Area" }, new[] { area }, new[] { "ID", "Image" });
            }
            else
            {
                tableImage = _dBContext.LoadDataTable(_TABLE_NAME + "_IMG_NAS", new[] { "Area" }, new[] { area }, null);
                DataTable dtZ = TDMK_ConverterService.JsonToDataTable(tableImage.Rows[0]["Data"].ToString());
                NasRepository _nas = new NasRepository();
                _nas.MergeDataTable(dtZ, _TABLE_NAME, itemCode, lotNo, tableImage.Rows[0]["LocationImg"].ToString());
                tableImage = dtZ;
                //_nas.MergeDataTable(, "" );
            }
            before = ConvertDataTable(TDMK_ConverterService.JsonToDataTable(json[1]), tableImage);

            Dictionary<string, Dictionary<string, DataTable>> dic = new Dictionary<string, Dictionary<string, DataTable>>();
            for (int i = 2; i < json.Count(); i++)
            {
                string[] bson = json[i].Replace("<", "").Replace(">", "").Split('%');
                string tape = bson[0];
                string name = bson[1];
                DataTable item = ConvertDataTable(TDMK_ConverterService.JsonToDataTable(bson[2]), tableImage);
                if (dic.TryGetValue(tape, out Dictionary<string, DataTable> jtem))
                {
                    jtem.Add(name, item);
                }
                else
                {
                    Dictionary<string, DataTable> zdic = new Dictionary<string, DataTable>();
                    zdic.Add(name, item);
                    dic.Add(tape, zdic);
                }
            }


            return dic;
        }
        public static string SetupSpec(ExcelWorksheet worksheet)
        {
            string spec = "";
            ExportProcess exportProcess = new ExportProcess();
            string[] healder = new[] { "Sample", "Sumitomo internal spec", "after ORT test" };
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(worksheet, healder, false);
            int bigNum = int.MinValue;
            if (dic.TryGetValue("Sample", out string address))
            {
                foreach (string item in address.Split('-'))
                {
                    string value = worksheet.Cells[item].Text.Replace("Sample", "").Trim();
                    if (int.TryParse(value, out int res))
                    {
                        bigNum = Math.Max(res, bigNum);
                    }
                }
            }
            if (bigNum != int.MinValue)
            {
                spec += $"{bigNum}:";
            }
            else
            {
                spec = "0:";
                Debugger.Break();
                // Dừng ở đây thì có lỗi xảy ra
            }
            Debugger.Break();

            if (dic.TryGetValue("Sumitomo internal spec", out address))
            {
                for (int i = 0; i < address.Split('-').Count(); i++)
                {
                    string add = address.Split('-')[i];
                    //tìm kiếm ort test gần nhất
                    int minHere = int.MaxValue;
                    string test = "";
                    foreach (var item in dic["after ORT test"].Split('-'))
                    {
                        int min = ExportProcess.DistanceRow(item, add);
                        if (min < minHere && min > 0)
                        {
                            minHere = min;
                            test = worksheet.Cells[item].Text.Trim();
                        }
                    }
                    spec += $"{test}!";
                    while (true)
                    {
                        add = ExportProcess.AddColumn(add, 1);
                        string value = worksheet.Cells[add].Text;
                        if (string.IsNullOrEmpty(value))
                        {
                            break;
                        }
                        string key = worksheet.Cells[ExportProcess.AddRow(add, -1)].Text;
                        spec += $"{key}={value}:";
                    }
                    if (i + 1 < address.Split('-').Count())
                    {
                        spec += "&";
                    }
                }
            }

            // SetupSpec method implementation
            return spec;
        }
        private void FillDataInTape(ExcelWorksheet worksheet, DataTable dataTable, DataTable spec, string address, int sample, string name, string tape)
        {
            string addressImageSample = ExportProcess.getRangeBaseAddressByCellAddress(worksheet, ExportProcess.AddRow(address, 1));
            byte[] imgBck = (byte[])((DataRow)spec.AsEnumerable().FirstOrDefault(row => row.Field<string>("Name") == name && row.Field<string>("TAPE") == tape))["Image Sample"];
            ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressImageSample], imgBck, $"BACK{tape}-{name}");
            int id = 0;
            string sampleAddress = ExportProcess.AddColumn(address, 2);
            foreach (DataRow dataRow in dataTable.Rows)
            {
                if (id >= sample)
                {
                    break;
                }
                string image = ExportProcess.AddRow(sampleAddress, 1);
                ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[image], (byte[])dataRow["Image"], $"Image{tape}-{name}-{id}");
                string graph = ExportProcess.AddRow(image, 1);
                ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[graph], (byte[])dataRow["Graph"], $"Graph{tape}-{name}-{id}");
                string productID = ExportProcess.AddRow(image, -2);
                try
                {

                    worksheet.Cells[productID].Value = dataRow["ProductID"];
                }
                catch
                {

                }
                sampleAddress = ExportProcess.AddColumn(sampleAddress, 1);
                string peak = ExportProcess.AddRow(graph, 1);
                worksheet.Cells[peak].Value = double.Parse(dataRow["Peak"].ToString());
                string avz = ExportProcess.AddRow(peak, 1);
                worksheet.Cells[avz].Value = double.Parse(dataRow["Average"].ToString());
                if (name.Equals("LINER"))
                {
                    avz = ExportProcess.AddRow(avz, 1);
                    worksheet.Cells[avz].Value = double.Parse(dataRow["Peak"].ToString()) * 0.0098;
                    avz = ExportProcess.AddRow(avz, 1);
                    worksheet.Cells[avz].Value = double.Parse(dataRow["Average"].ToString()) * 0.0098;
                }
                string az = ExportProcess.AddRow(avz, 1);

                worksheet.Cells[az].Value = dataRow["Judgement Peeling force"];
                az = ExportProcess.AddRow(az, 1);
                worksheet.Cells[az].Value = dataRow["Judgement failure mode"];
                id++;

            }
        }

        public string Export(string itemCode, string lotNo)
        {
            DataTable dataTable = new DataTable();
            Dictionary<string, Dictionary<string, DataTable>> dic = Load(itemCode, lotNo, ref dataTable);
            if (dataTable.Rows.Count < 0)
            {
                return "Không có dữ liệu của itemcode lotno";
            }
            DataTable spec = _dBContext.LoadDataTable("SPEC_COMMENT_3", new[] { "ItemCode", "Sheet" }, new[] { itemCode, "ENVIRONMENT_EN-DURANCE" });
            if (spec.Rows.Count <= 0)
            {
                return "Spec chưa được cài đặt";
            }
            int sample = int.Parse(spec.Rows[0]["Count_Sample"].ToString());
            ExportProcess exportProcess = new ExportProcess();
            using (ExcelPackage ex = exportProcess.FindFormatProcess(FORMAT_NAME, itemCode, lotNo))
            {
                using (ExcelWorksheet workSheet = exportProcess.FindSheet(ex, FORMAT_NAME))
                {
                    string sampleSTR = $"Sample {sample}";
                    string[] colHeader = new[] { "Liner peeling after ORT test", "PSA peeling after ORT test", sampleSTR, "Flex SN", "Average Force" };
                    IDictionary<string, string> dicHeader = ExportProcess.FindAddressByText(workSheet, colHeader);


                    Dictionary<string, string> dicCol = new Dictionary<string, string>();

                    #region Find area PSA Lineer
                    foreach (string item in new[] { "Liner peeling after ORT test", "PSA peeling after ORT test" })
                    {
                        dicCol.Add(item + "Flex SN", dicHeader[item]);
                        if (dicHeader.TryGetValue("Flex SN", out string value) && dicHeader.TryGetValue(item, out string valueZ))
                        {
                            string[] flexCout = value.Split('-');
                            int min = int.MaxValue;
                            foreach (var item1 in flexCout)
                            {
                                int z = ExportProcess.DistanceRow(valueZ, item1);
                                if (z < min && z >= 0)
                                {
                                    dicCol[item + "Flex SN"] = item1;
                                    min = z;
                                }
                            }
                            if (min == int.MaxValue)
                            {
                                return "Lỗi về lấy flex sn";
                            }

                        }
                        else
                        {
                            return "Không có đủ flexSN";
                        }
                        dicCol.Add(item + "Average Force", dicHeader[item]);
                        if (dicHeader.TryGetValue("Average Force", out value) && dicHeader.TryGetValue(item, out valueZ))
                        {
                            string[] flexCout = value.Split('-');
                            int min = int.MaxValue;
                            foreach (var item1 in flexCout)
                            {
                                int z = ExportProcess.DistanceRow(valueZ, item1);
                                if (z < min && z >= 0)
                                {
                                    dicCol[item + "Average Force"] = item1;
                                    min = z;
                                }
                            }
                            if (min == int.MaxValue)
                            {
                                return "Lỗi về lấy Average Force";
                            }

                        }
                        else
                        {
                            return "Không có đủ Average Force";
                        }
                        dicCol.Add(item + sampleSTR, dicHeader[item]);
                        if (dicHeader.TryGetValue(sampleSTR, out value) && dicHeader.TryGetValue(item, out valueZ))
                        {
                            string[] flexCout = value.Split('-');
                            int min = int.MaxValue;
                            foreach (var item1 in flexCout)
                            {
                                int z = ExportProcess.DistanceRow(valueZ, item1);
                                if (z < min && z >= 0)
                                {
                                    dicCol[item + sampleSTR] = item1;
                                    min = z;
                                }
                            }
                            if (min == int.MaxValue)
                            {
                                return "Lỗi về lấy sampleSTR";
                            }

                        }
                        else
                        {
                            return "Không có đủ sampleSTR";
                        }

                    }
                    #endregion

                    #region Copy and paste
                    string[] countPSA = dataTable.AsEnumerable().Where(row => row.Field<string>("name") == "PSA").Select(row => row.Field<string>("TAPE")).ToArray();
                    string[] countLiner = dataTable.AsEnumerable().Where(row => row.Field<string>("name") == "LINER").Select(row => row.Field<string>("TAPE")).ToArray();

                    if (dicCol.TryGetValue("PSA peeling after ORT testFlex SN", out string addressFlexSN)
                        && dicCol.TryGetValue($"PSA peeling after ORT test{sampleSTR}", out string addressSample)
                        && dicCol.TryGetValue("PSA peeling after ORT testAverage Force", out string addressCPK))
                    {
                        string addressPointer = workSheet.Cells[workSheet.Cells[addressCPK].Start.Row + 1, workSheet.Cells[addressFlexSN].Start.Column].Address;

                        for (int i = 0; i < countPSA.Count() - 1; i++)
                        {
                            string point = addressPointer;
                            if (i == 0)
                            {
                                workSheet.Cells[ExportProcess.AddRow(addressFlexSN, 1)].Value = countPSA[countPSA.Count() - 1] + "%PSA";
                            }
                            string addressRange = $"{addressFlexSN}:{workSheet.Cells[workSheet.Cells[addressCPK].Start.Row, workSheet.Cells[addressSample].Start.Column].Address}";
                            exportProcess.CopyAndInsert(workSheet, addressRange, ref addressPointer, true);
                            workSheet.Cells[ExportProcess.AddRow(point, 2)].Value = countPSA[i] + "%PSA";

                        }
                    }
                    if (dicCol.TryGetValue("Liner peeling after ORT testFlex SN", out addressFlexSN)
                        && dicCol.TryGetValue($"Liner peeling after ORT test{sampleSTR}", out addressSample)
                        && dicCol.TryGetValue("Liner peeling after ORT testAverage Force", out addressCPK))
                    {
                        string addressPointer = workSheet.Cells[workSheet.Cells[addressCPK].Start.Row + 1, workSheet.Cells[addressFlexSN].Start.Column].Address;

                        for (int i = 0; i < countLiner.Count() - 1; i++)
                        {
                            string point = addressPointer;
                            if (i == 0)
                            {
                                workSheet.Cells[ExportProcess.AddRow(addressFlexSN, 1)].Value = countLiner[countLiner.Count() - 1] + "%LINER";
                            }
                            string addressRange = $"{addressFlexSN}:{workSheet.Cells[workSheet.Cells[addressCPK].Start.Row, workSheet.Cells[addressSample].Start.Column].Address}";
                            exportProcess.CopyAndInsert(workSheet, addressRange, ref addressPointer, true);
                            workSheet.Cells[ExportProcess.AddRow(point, 2)].Value = countLiner[i] + "%LINER";

                        }
                    }

                    #endregion

                    #region Fill data
                    List<string> tapeList = new List<string>();
                    foreach (string item in countPSA)
                    {
                        tapeList.Add(item + "%PSA");
                    }
                    foreach (string item in countLiner)
                    {
                        tapeList.Add(item + "%LINER");
                    }
                    IDictionary<string, string> dicTape = ExportProcess.FindAddressByText(workSheet, tapeList.ToArray());

                    foreach (string item in dicTape.Keys)
                    {
                        string tape = item.Split('%')[0];
                        string name = item.Split('%')[1];
                        if ((dic[tape]).TryGetValue(name, out DataTable dt))
                        {
                            FillDataInTape(workSheet, dt, dataTable, dicTape[item], sample, name, tape);
                        }
                    }
                    #endregion
                    
                    exportProcess.SaveExcelWorksheet(ex, FORMAT_NAME, $"{itemCode}_{lotNo}", "NPI", false);
                }

            }

            return "OK";
        }

    }

}

