
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OK2SHIP_SMT.Libary;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZedGraph;
using static System.Resources.ResXFileRef;

namespace OK2SHIP_SMT.Services
{
    public class EEDService
    {

        private ExportProcess _export = new ExportProcess();
        private DBContext _dBContext = new DBContext();
        private const string FORMAT_NAME = "Environment en-durance";
        private const string _TABLE_NAME = "ENVIRONMENT_EN_DURANCE";
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
                    exportProcess.SaveExcelWorksheet(ex, FORMAT_NAME, itemCode, lotNo);
                }

            }

            return "Export thành công";
        }
        private void FillDataInTape(ExcelWorksheet worksheet, DataTable dataTable, DataTable spec, string address, int sample, string name, string tape)
        {
            string addressImageSample = ExportProcess.getRangeBaseAddressByCellAddress(worksheet, ExportProcess.AddRow(address, 1));
            byte[] imgBck = TDMK_ImageConverter.ImageToByteArray((Image)((DataRow)spec.AsEnumerable().FirstOrDefault(row => row.Field<string>("Name") == name && row.Field<string>("TAPE") == tape))["Image Sample"], ImageFormat.Jpeg);
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

                ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[image], TDMK_ImageConverter.ImageToByteArray((Image)dataRow["Image"], ImageFormat.Jpeg), $"Image{tape}-{name}-{id}");
                string graph = ExportProcess.AddRow(image, 1);
                ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[graph], TDMK_ImageConverter.ImageToByteArray((Image)dataRow["Graph"], ImageFormat.Jpeg), $"Graph{tape}-{name}-{id}");
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
        public bool CheckLoaction(string location, string itemCode, string lotNo)
        {
            location = FileFolderRepository.GetFolderName(location);
            string[] array = location.Split(new[] { '-', '_' });
            string itemCodeL = $"{array[1]}-{array[2]}";
            try
            {
                itemCodeL += $"-{ValidateService.isDigit(array[3])}";
            }
            catch
            {
                // không có lotcut
            }

            return ValidateService.compareItemCodeLotNo(itemCodeL, $"{itemCode}-{lotNo}"); ;
        }
        public string[] ReadFile(string location, string itemCode, string lotNo)
        {
            location = location.Trim();
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            if (!CheckLoaction(location, itemCode, lotNo))
            {
                throw new Exception("Không trùng itemCode, lotno");
            }
            if (!FileFolderRepository.checkLocationIsValid(location))
            {
                throw new Exception("Đường dẫn không hợp lệ");
            }
            string[] subLocation = FileFolderRepository.GetSubFolders(location);
            List<string> result = new List<string>();
            foreach (var item in subLocation)
            {
                string process = FileFolderRepository.GetFolderName(item).Replace(" ", "");
                switch (process)
                {
                    case "LINER":
                        result.Add(item);
                        break;

                    case "PSA":
                        result.Add(item);
                        break;
                    default:
                        break;
                }
            }
            return result.ToArray();

        }
        private DataTable SolveLinerPSAFolder(string location, string process)
        {
            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("PCS", typeof(string));
            dataTable.Columns.Add("Image", typeof(Image));
            dataTable.Columns.Add("Graph", typeof(Image));
            dataTable.Columns.Add("Tape", typeof(string));
            if (process.Contains("LINER"))
            {
                dataTable.Columns.Add("Peak(Gf)", typeof(string));
                dataTable.Columns.Add("Average(Gf)", typeof(string));
            }
            dataTable.Columns.Add("Peak", typeof(string));
            dataTable.Columns.Add("Average", typeof(string));
            dataTable.Columns.Add("Judgement Peeling force", typeof(string));
            dataTable.Columns.Add("Judgement failure mode", typeof(string));

            string[] subFolders = FileFolderRepository.GetSubFolders(location);
            foreach (var item in subFolders)
            {
                string tape = FileFolderRepository.GetFolderName(item);
                if (tape.Contains("TAPE"))
                {
                    IList<KeyValuePair<Image, string>> listImage = FileFolderRepository.ListAllPictureInAFolder(item);
                    var orderList = listImage.OrderBy(x => int.Parse(x.Value.Split('.')[0]));
                    foreach (var itemImage in orderList)
                    {
                        DataRow row = dataTable.NewRow();
                        string numPcs = itemImage.Value.Split('.')[0].Trim();
                        row["PCS"] = numPcs;
                        row["Image"] = itemImage.Key;
                        row["Tape"] = tape;
                        row["Graph"] = SolveFileLinerPSQ($"{item}\\{numPcs}.xlsx", out string resFile);
                        if (process.Contains("LINER"))
                        {
                            row["Peak(Gf)"] = resFile.Split(':')[0];
                            row["Average(Gf)"] = resFile.Split(':')[1];
                            row["Peak"] = double.Parse(resFile.Split(':')[0]) * 0.0098;
                            row["Average"] = double.Parse(resFile.Split(':')[1]) * 0.0098;
                        }
                        else
                        {
                            row["Peak"] = resFile.Split(':')[0];
                            row["Average"] = resFile.Split(':')[1];
                        }
                        dataTable.Rows.Add(row);
                    }
                }
            }

            return dataTable;
        }
        private Image SolveFileLinerPSQ(string localtion, out string graph)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(localtion))
                {
                    using (ExcelWorksheet workSheet = package.Workbook.Worksheets[0])
                    {
                        IDictionary<string, string> dic = ExportProcess.FindAddressByText(workSheet, new[] { "Max", "Average" }, true);
                        if (!dic.TryGetValue("Max", out string max))
                        {
                            max = "";
                        }
                        if (!dic.TryGetValue("Average", out string min))
                        {
                            min = "";
                        }
                        max = workSheet.Cells[ExportProcess.AddColumn(max, 4)].Text;
                        min = workSheet.Cells[ExportProcess.AddColumn(min, 4)].Text;
                        graph = $"{max.Trim()}:{min.Trim()}";
                        ExcelPicture pic = workSheet.Drawings["Picture 1"] as ExcelPicture;
                        return TDMK_ImageConverter.ByteArrayToImage(pic.Image.ImageBytes);
                    }
                }
            }
            catch
            {
                throw new Exception("error");
                //return ;
            }
        }
        public static List<DataTable> SplitDataTableByTape(DataTable sourceTable, int pcs)
        {
            if (!sourceTable.Columns.Contains("tape"))
            {
                throw new ArgumentException("DataTable không chứa cột 'Tape'.");
            }

            List<DataTable> resultTables = new List<DataTable>();

            // 1. Lấy các giá trị duy nhất từ cột "tape"
            var distinctTapeValues = sourceTable.AsEnumerable()
                                              .Select(row => row.Field<string>("Tape"))
                                              .Distinct()
                                              .ToList();

            // 2. Lặp qua từng giá trị duy nhất
            foreach (string tapeValue in distinctTapeValues)
            {
                // 3. Lọc DataTable ban đầu để tạo DataTable mới
                DataTable newTable = sourceTable.Clone(); // Sao chép cấu trúc của DataTable ban đầu

                foreach (DataRow row in sourceTable.Rows)
                {
                    if (row.Field<string>("Tape") == tapeValue)
                    {
                        newTable.ImportRow(row); // Sao chép các hàng thỏa mãn điều kiện
                    }
                }
                if (pcs < 0)
                {
                    resultTables.Add(newTable);

                }
                else
                {
                    resultTables.Add(newTable.AsEnumerable().Take(pcs).CopyToDataTable());
                }
            }

            return resultTables;
        }
        private const string _SAMPLE_DIC = "SAMPLE";
        public void SolveFolderPSALiner(string location, Dictionary<string, Dictionary<string, DataTable>> dic, DataTable spec, string processz, int pcs = -1)
        {

            if (spec.Columns.Count <= 0)
            {
                spec.Columns.Add("TAPE");
                spec.Columns.Add("Name");
                spec.Columns.Add("Peak Peeling Force (N)");
                spec.Columns.Add("Average Peeling Force (N)");
                spec.Columns.Add($"Image Sample", typeof(Image));

            }
            string result = FileFolderRepository.GetFolderName(location).Replace(" ", "");
            DataTable dataTableLiner = SolveLinerPSAFolder(location, processz);
            IList<DataTable> list = SplitDataTableByTape(dataTableLiner, pcs);

            DataTable sample = new DataTable();
            sample.Columns.Add("STT");
            sample.Columns.Add("Name");
            sample.Columns.Add("Tape");
            sample.Columns.Add("Image", typeof(Image));
            sample.Columns.Add("Graph", typeof(Image));



            foreach (var item in list)
            {
                string tape = item.Rows[0]["Tape"].ToString();
                if (!dic.ContainsKey(tape))
                {
                    Dictionary<string, DataTable> dicTemp = new Dictionary<string, DataTable>();
                    dicTemp.Add(result, item);
                    dic.Add(tape, dicTemp);

                }
                else
                {
                    if (!dic[tape].ContainsKey(result))
                    {
                        dic[tape].Add(result, item);
                    }
                    else
                    {
                        dic[tape][result].Merge(item);
                    }
                }

                DataRow row = spec.NewRow();
                row["TAPE"] = tape;
                row["Name"] = result;
                row["Image Sample"] = ((DataTable)dic[tape][result]).Rows[0]["Image"];
                spec.Rows.Add(row);

            }
        }
        public void SolveFolder(string location, out string result)
        {
            result = FileFolderRepository.GetFolderName(location).Replace(" ", "");
            switch (result)
            {
                case "LINER":
                case "PSA":
                    return;
                default:
                    throw new Exception("Không có folder này");
            }

            throw new Exception("Folder này không có dữ liệu khớp");
        }
        public DataTable getDTImgStructor()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Area", typeof(int));
            dt.Columns.Add("Image", typeof(Image));
            return dt;
        }
        public int save(string itemCode, string lotNo, DataTable before, Dictionary<string, Dictionary<string, DataTable>> dic, int prime = -1)
        {
            int res = 0;
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            int id = prime;
            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
            {
                throw new Exception("ItemCode or LotNo is null");
            }
            if (prime == -1)
            {
                DataTable checker = _dBContext.LoadDataTable(_TABLE_NAME, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo }, new[] { "ID", "Data" });
                if (checker.Rows.Count > 0)
                {
                    DataRow row = checker.Rows[0];
                    throw new Exception($"1234 - {row["Data"].ToString().Split('#')[0]} - ItemCode and LotNo already exist");
                }
            }

            id = _dBContext.GetID(_TABLE_NAME + "_IMG") + 1;

            DataTable imageDT = getDTImgStructor();
            int area = id;
            #region before
            string list = $"{area.ToString()}#" + ConvertDataTable(before, imageDT, ref id, area);
            #endregion
            #region psa linear
            foreach (string keyOut in dic.Keys)
            {
                foreach (string keyIn in dic[keyOut].Keys)
                {
                    DataTable dt = dic[keyOut][keyIn];
                    list += "#" + $"<{keyOut}%{keyIn}%{ConvertDataTable(dt, imageDT, ref id, area)}>";
                }
            }
            #endregion
            #region Insert environment
            DataTable DATA = _dBContext.GetTableStructure(_TABLE_NAME);
            DataRow rowz = DATA.NewRow();
            rowz["ItemCode"] = itemCode;
            rowz["LotNo"] = lotNo;
            rowz["Data"] = list;
            DATA.Rows.Add(rowz);
            NasRepository _nas = new NasRepository();
            string location = _nas.HandleImageDataTable(imageDT, $"{_TABLE_NAME}", itemCode, lotNo);
            string Data = ConverterService.DataTableToJson(imageDT);
            DataTable imgDT = _dBContext.GetTableStructure($"{_TABLE_NAME}_IMG_NAS");
            DataRow roDT = imgDT.NewRow();
            roDT["Area"] = area;
            roDT["Data"] = Data;
            roDT["LocationImg"] = location;
            imgDT.Rows.Add(roDT);
            res += _dBContext.BuckDataTable(imgDT, $"{_TABLE_NAME}_IMG_NAS", new[] { "Area" }, null, "ID");
            //res += _dBContext.BuckDataTable(imageDT, _TABLE_NAME + "_IMG", new[] { "Area" });
            res += _dBContext.BuckDataTable(DATA, _TABLE_NAME, new[] { "ItemCode", "LotNo" }, null, "ID");
            #endregion
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
                        Image image = (Image)row[colName];
                        if (datatable.Columns[colName].DataType.FullName != "System.Drawing.Image")
                        {
                            image = TDMK_ImageConverter.ByteArrayToImage((byte[])row[colName]);

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
                    res.Columns.Add(col.ColumnName.Replace("&CONVERTER", ""), typeof(Image));
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
                        Image img = (Image)dataTableImage.Rows[point]["Image"];
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
                DataTable dtZ = ConverterService.JsonToDataTable(tableImage.Rows[0]["Data"].ToString());
                NasRepository _nas = new NasRepository();
                _nas.MergeDataTable(dtZ, _TABLE_NAME, itemCode, lotNo, tableImage.Rows[0]["LocationImg"].ToString());
                tableImage = dtZ;
                //_nas.MergeDataTable(, "" );
            }
            before = ConvertDataTable(ConverterService.JsonToDataTable(json[1]), tableImage);

            Dictionary<string, Dictionary<string, DataTable>> dic = new Dictionary<string, Dictionary<string, DataTable>>();
            for (int i = 2; i < json.Count(); i++)
            {
                string[] bson = json[i].Replace("<", "").Replace(">", "").Split('%');
                string tape = bson[0];
                string name = bson[1];
                DataTable item = ConvertDataTable(ConverterService.JsonToDataTable(bson[2]), tableImage);
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

        public void FillProductID(Dictionary<string, Dictionary<string, DataTable>> dic, string itemCode, string lotNo, string location)
        {
            if (string.IsNullOrEmpty(location.Trim()))
            {
                throw new Exception("Điền đường dẫn productID");
            }
            ProductIDService productIDService = new ProductIDService(itemCode, lotNo, location, new[] { "ORT", "environment" }, new[] { "LINER", "PSA" });
            Dictionary<string, string> diczz = productIDService._listFile;

            foreach (Dictionary<string, DataTable> dicz in dic.Values)
            {
                foreach (var item in dicz)
                {
                    if (diczz.TryGetValue(item.Key, out string locationz))
                    {
                        List<string> _PRODUCT_ID = productIDService.getListProductID(locationz);
                        item.Value.Columns.Add("ProductID");
                        foreach (DataRow row in item.Value.Rows)
                        {
                            if (_PRODUCT_ID.Count > 0)
                            {
                                row["ProductID"] = _PRODUCT_ID[0];
                                _PRODUCT_ID.Remove(_PRODUCT_ID[0]);
                            }

                        }
                    }
                }
            }
        }

        public DataTable GetSpec(string itemCode)
        {
            itemCode = itemCode.Trim();
            DataTable dataTable = _dBContext.LoadDataTable("SPEC_COMMENT_3", new[] { "ItemCode", "Sheet" }, new[] { itemCode, "ENVIRONMENT_EN-DURANCE" });
            return dataTable;
        }

        public void FillSpec(DataTable spec_log, DataTable spec)
        {
            string content = spec_log.Rows[0]["Location"].ToString();
            foreach (var item in content.Split('&'))
            {
                string type = item.Split('!')[0].ToUpper().Contains("PSA") ? "PSA" : "LINER";
                string[] containz = item.Split('!')[1].Split(':');
                foreach (DataRow row in spec.Rows)
                {
                    if (row["Name"].ToString().Contains(type))
                    {
                        foreach (string item1 in containz)
                        {
                            string healder = item1.Split('-')[0].Replace(" ", "").ToUpper();
                            if (healder.Contains("PEAKPEELINGFORCE(N)"))
                            {
                                row["Peak Peeling Force (N)"] = item1.Split('=')[1];
                            }
                            if (healder.Contains("AVERAGEPEELINGFORCE(N)"))
                            {
                                row["Average Peeling Force (N)"] = item1.Split('=')[1];
                            }
                        }
                    }
                }
            }
        }

        public void CheckSpec(DataTable spec, Dictionary<string, Dictionary<string, DataTable>> dic)
        {

            foreach (string tape in dic.Keys)
            {
                foreach (string name in dic[tape].Keys)
                {
                    string avarage = "";
                    string peak = "";
                    foreach (DataRow row in spec.Rows)
                    {
                        if (row["Name"].Equals(name))
                        {
                            if (row["TAPE"].Equals(tape))
                            {
                                avarage = row["Peak Peeling Force (N)"].ToString();
                                peak = row["Average Peeling Force (N)"].ToString();
                            }

                        }
                        if (!string.IsNullOrEmpty(avarage) && !string.IsNullOrEmpty(peak))
                        {
                            double aS = double.Parse(avarage.Split('~')[0]);
                            double aE = double.Parse(avarage.Split('~')[1]);
                            double pS = double.Parse(peak.Split('~')[0]);
                            double pE = double.Parse(peak.Split('~')[1]);
                            foreach (DataRow rowz in dic[tape][name].Rows)
                            {
                                bool prime = false;
                                double peakR = double.Parse(rowz["Peak"].ToString());
                                double averageR = double.Parse(rowz["Average"].ToString());

                                prime = peakR < pE && peakR > pS;
                                prime = prime && averageR < aE && averageR > aS;
                                rowz["Judgement Peeling force"] = prime ? "Pass" : "Fail";
                            }
                        }
                    }
                }
            }
        }
    }
}
