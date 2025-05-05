
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OK2SHIP_SMT.Libary;
using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZedGraph;

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
            Dictionary<string, Dictionary<string, DataTable>> dic = Load(itemCode, lotNo, out dataTable);
            if (dataTable.Rows.Count < 0)
            {
                return "Không có dữ liệu của itemcode lotno";
            }

            ExportProcess exportProcess = new ExportProcess();
            using (ExcelPackage ex = exportProcess.FindFormatProcess(FORMAT_NAME, itemCode, lotNo))
            {
                using (ExcelWorksheet workSheet = exportProcess.FindSheet(ex, FORMAT_NAME))
                {
                    string[] strFind = new[] { "Sample 32", "Flex SN", "CPK", "Liner peeling after ORT test", "PSA peeling after ORT test" };
                    IDictionary<string, string> addDic = ExportProcess.FindAddressByText(workSheet, strFind, false);
                    int numCopy = dic.Keys.Count - 1;
                    string linerAdd = "";
                    string PSAAdd = "";
                    for (int i = 0; i < 2; i++)
                    {
                        if (addDic.TryGetValue("Sample 32", out string addStart) && addDic.TryGetValue("CPK", out string cpk))
                        {
                            addStart = addStart.Split('-')[i];
                            cpk = cpk.Split('-')[i];
                            addStart = workSheet.Cells[workSheet.Cells[cpk].End.Row, workSheet.Cells[addStart].End.Column].Address;
                        }
                        else
                        {
                            throw new Exception("Lost data");
                        }
                        if (addDic.TryGetValue("Flex SN", out string addEnd))
                        {
                            addEnd = addEnd.Split('-')[i];
                        }
                        else
                        {
                            throw new Exception("Lost data");
                        }
                        switch (i + 1)
                        {
                            case 1:
                                linerAdd = $"{addEnd}:{addStart}";
                                break;
                            case 2:
                                PSAAdd = $"{addEnd}:{addStart}";
                                break;
                            default:
                                throw new Exception("Lost data");
                        }
                    }
                    //PSA
                    string addressPointer = workSheet.Cells[workSheet.Cells[ExportProcess.AddRow(PSAAdd.Split(':')[1], 2)].End.Row, workSheet.Cells[ExportProcess.AddRow(PSAAdd.Split(':')[0], 2)].End.Column].Address;
                    workSheet.Cells[ExportProcess.AddRow(PSAAdd.Split(':')[0], 1)].Value = dic.Keys.ToArray()[1];
                    for (int i = 0; i < numCopy - 1; i++)
                    {
                        string point = addressPointer;
                        exportProcess.CopyAndInsert(workSheet, PSAAdd, ref addressPointer, false);
                        addressPointer = ExportProcess.AddRow(addressPointer, 2);
                        workSheet.Cells[ExportProcess.AddRow(point, 1)].Value = dic.Keys.ToArray()[i + 2];
                    }

                    addressPointer = workSheet.Cells[workSheet.Cells[ExportProcess.AddRow(linerAdd.Split(':')[1], 2)].End.Row, workSheet.Cells[ExportProcess.AddRow(linerAdd.Split(':')[0], 2)].End.Column].Address;
                    workSheet.Cells[ExportProcess.AddRow(linerAdd.Split(':')[0], 1)].Value = dic.Keys.ToArray()[1];
                    for (int i = 0; i < numCopy - 1; i++)
                    {
                        string point = addressPointer;
                        exportProcess.CopyAndInsert(workSheet, linerAdd, ref addressPointer, true);
                        addressPointer = ExportProcess.AddRow(addressPointer, 2);
                        workSheet.Cells[ExportProcess.AddRow(point, 1)].Value = dic.Keys.ToArray()[i + 2];
                    }
                    IDictionary<string, string> dicZZ = ExportProcess.FindAddressByText(workSheet, dic.Keys.ToArray());
                    //DataTable beforeImage = dic["SAMPLE"]["SAMPLE"];
                    foreach (string tape in dic.Keys)
                    {
                        if (!tape.Contains("SAMPLE"))
                        {
                            if (dic.TryGetValue(tape, out Dictionary<string, DataTable> valueax))
                            {
                                foreach (string area in valueax.Keys)
                                {
                                    DataTable data = valueax[area];
                                    int index = area.Equals("LINER") ? 0 : area.Equals("PSA") ? 1 : -1;
                                    if (dicZZ.TryGetValue(tape, out string address))
                                    {
                                        address = address.Split('-')[index];
                                        address = ExportProcess.AddRow(ExportProcess.AddColumn(address, 2), 1);
                                        if (index == 0)
                                        {
                                            for (int i = 0; i < dataTable.Columns.Count; i++)
                                            {
                                                exportProcess.InsertImageToCell(workSheet, workSheet.Cells[ExportProcess.AddColumn(address, i)], (byte[])dataTable.Rows[0][i], $"beforeImage{i}{area}{tape}");
                                                exportProcess.InsertImageToCell(workSheet, workSheet.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 1), i)], (byte[])data.Rows[i]["Image"], $"after{i}{area}{tape}");
                                                exportProcess.InsertImageToCell(workSheet, workSheet.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 2), i)], (byte[])data.Rows[i]["Graph"], $"graph{i}{area}{tape}");
                                                if (double.TryParse(data.Rows[i]["Peak"].ToString(), out double peak) && double.TryParse(data.Rows[i]["Average"].ToString(), out double average))
                                                {
                                                    workSheet.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 3), i)].Value = peak;
                                                    workSheet.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 4), i)].Value = average;
                                                    workSheet.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 5), i)].Value = peak * 1000;
                                                    workSheet.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 6), i)].Value = average * 1000;
                                                }
                                                workSheet.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 7), i)].Value = data.Rows[i]["Judgement Peeling force"].ToString();
                                                workSheet.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 8), i)].Value = data.Rows[i]["Judgement failure mode"].ToString();

                                            }
                                        }
                                        else
                                        {
                                            for (int i = 0; i < dataTable.Columns.Count; i++)
                                            {
                                                exportProcess.InsertImageToCell(workSheet, workSheet.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 0), i)], (byte[])data.Rows[i]["Image"], $"imagePSA{i}{area}{tape}");
                                                exportProcess.InsertImageToCell(workSheet, workSheet.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 1), i)], (byte[])data.Rows[i]["Graph"], $"graphPSA{i}{area}{tape}");
                                                workSheet.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 2), i)].Value = data.Rows[i]["Peak"];
                                                workSheet.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 3), i)].Value = data.Rows[i]["Average"];
                                                workSheet.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 4), i)].Value = data.Rows[i]["Judgement Peeling force"];
                                                workSheet.Cells[ExportProcess.AddColumn(ExportProcess.AddRow(address, 5), i)].Value = data.Rows[i]["Judgement failure mode"];
                                            }
                                        }
                                    }


                                    //Debugger.Break();
                                }
                            }
                        }
                        else
                        {
                            Dictionary<string, DataTable> listSample = dic["SAMPLE"];
                            foreach (string range in listSample.Keys)
                            {
                                int index = range.Equals("LINER") ? 0 : range.Equals("PSA") ? 1 : -1;
                                foreach (DataRow row in listSample[range].Rows)
                                {
                                    string tapez = row["Tape"].ToString();
                                    if (dicZZ.TryGetValue(tapez, out string address))
                                    {
                                        address = ExportProcess.AddRow(address.Split('-')[index], 1);
                                        address = exportProcess.getRangeBaseAddressByCellAddress(workSheet, address);
                                        exportProcess.InsertImageToCell(workSheet, workSheet.Cells[address], (byte[])row["Image"], $"sample{tapez}{range}");
                                        //Debugger.Break();
                                    }
                                }

                            }
                        }
                    }

                    exportProcess.SaveExcelWorksheet(ex, FORMAT_NAME, itemCode, lotNo);
                }
            }
            return "Export thành công";
        }
        public bool CheckLoaction(string location, string itemCode, string lotNo)
        {
            location = FileFolderRepository.GetFolderName(location);
            string[] array = location.Split('-');
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
                    case "LINERTRUOCKEO":
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
        private DataTable SolveLinerPSAFolder(string location)
        {
            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("PCS", typeof(string));
            dataTable.Columns.Add("Image", typeof(Image));
            dataTable.Columns.Add("Graph", typeof(Image));
            dataTable.Columns.Add("Tape", typeof(string));
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
                        row["Peak"] = resFile.Split(':')[0];
                        row["Peak"] = resFile.Split(':')[0];
                        row["Average"] = resFile.Split(':')[1];
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
        private DataTable SolveBeforeFolder(string location)
        {
            DataTable dataTable = new DataTable();

            dataTable.Columns.Add($"Image", typeof(Image));

            dataTable.Rows.Add();
            IList<KeyValuePair<Image, string>> listImg = FileFolderRepository.ListAllPictureInAFolder(location);
            var orderList = listImg.OrderBy(x => int.Parse(x.Value.Split('.')[0]));
            foreach (KeyValuePair<Image, string> item in orderList)
            {
                string stt = item.Value.Trim().Split('.')[0].Trim();
                dataTable.Columns.Add($"Image{stt}", typeof(Image));
                //ImageFormat format = ImageFormat.Jpeg;
                //string str = TDMK_ImageConverter.ImageToBase64((Image)item.Key, format);
                //Image image = TDMK_ImageConverter.Base64ToImage(str);
                //string strz = TDMK_ImageConverter.ImageToBase64(image, format);
                dataTable.Rows[0][$"Image{stt}"] = (Image)item.Key;
            }
            dataTable.Columns.Remove("Image");
            return dataTable;
        }
        public static List<DataTable> SplitDataTableByTape(DataTable sourceTable)
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

                resultTables.Add(newTable);
            }

            return resultTables;
        }
        private const string _SAMPLE_DIC = "SAMPLE";
        public void SolveFolderPSALiner(string location, Dictionary<string, Dictionary<string, DataTable>> dic)
        {

            string result = FileFolderRepository.GetFolderName(location).Replace(" ", "");
            DataTable dataTableLiner = SolveLinerPSAFolder(location);
            IList<DataTable> list = SplitDataTableByTape(dataTableLiner);

            DataTable sample = new DataTable();
            sample.Columns.Add("STT");
            sample.Columns.Add("Name");
            sample.Columns.Add("Tape");
            sample.Columns.Add("Image", typeof(Image));
            sample.Columns.Add("Graph", typeof(Image));

            if (!dic.ContainsKey(_SAMPLE_DIC))
            {
                dic.Add(_SAMPLE_DIC, new Dictionary<string, DataTable>());
            }


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

                if (dic.TryGetValue(_SAMPLE_DIC, out Dictionary<string, DataTable> dicz))
                {
                    string key = result;
                    if (!dicz.TryGetValue(key, out DataTable dtzs))
                    {
                        dicz.Add(key, sample);
                    }
                    if (!sample.AsEnumerable().Any(row => row.Field<string>("Tape") == tape))
                    {
                        DataRow row = sample.NewRow();
                        row["STT"] = sample.Rows.Count + 1;
                        row["Name"] = result;
                        row["Tape"] = tape;
                        row["Image"] = ((DataTable)dic[tape][result]).Rows[0]["Image"];
                        sample.Rows.Add(row);
                        //dic[_SAMPLE_DIC][key].Merge(row);
                    }

                }
            }
        }
        public DataTable SolveFolder(string location, out string result)
        {
            result = FileFolderRepository.GetFolderName(location).Replace(" ", "");
            switch (result)
            {
                case "LINERTRUOCKEO":
                    return SolveBeforeFolder(location);
                case "LINER":
                case "PSA":
                    return new DataTable();
                default:
                    throw new Exception("Không có folder này");
            }

            throw new Exception("Folder này không có dữ liệu khớp");
        }

        public int save(string itemCode, string lotNo, DataTable before, Dictionary<string, Dictionary<string, DataTable>> dic, bool prime = false)
        {
            int res = 0;
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            if (!prime)
            {
                if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
                {
                    throw new Exception("ItemCode or LotNo is null");
                }
                DataTable checker = _dBContext.LoadDataTable(_TABLE_NAME, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo }, new[] { "ID" });
                if (checker.Rows.Count > 0)
                {
                    throw new Exception("1234 - ItemCode and LotNo already exist");
                }
            }
            DataTable dataTableImage = _dBContext.GetTableStructure(_TABLE_NAME + "_IMAGE");
            int pcs = _dBContext.GetID(_TABLE_NAME + "_IMAGE") + 1;
            DataTable dataTable = _dBContext.GetTableStructure(_TABLE_NAME);
            #region Before
            string resData = "";
            DataRow rowBefore = dataTableImage.NewRow();
            for (int i = 1; i < dataTableImage.Columns.Count; i++)
            {
                var imz = before.Rows[0][before.Columns[i - 1].ColumnName];
                if (imz is Image)
                {
                    rowBefore[dataTableImage.Columns[i].ColumnName] = TDMK_ImageConverter.ImageToByteArray((Image)imz, ImageFormat.Jpeg);
                }
                else
                {
                    rowBefore[dataTableImage.Columns[i].ColumnName] = imz;
                }
            }
            rowBefore["ID"] = pcs;
            resData += $"Before<ImageID: {pcs++}>";
            dataTableImage.Rows.Add(rowBefore);
            #endregion

            DataRow dr = dataTable.NewRow();
            dr["ItemCode"] = itemCode;
            dr["LotNo"] = lotNo;
            ///Dictionary<string, Dictionary<string, DataTable>>
            /// TAPE - PSA/Linar - data
            foreach (string tape in dic.Keys)
            {
                Dictionary<string, DataTable> dicz = dic[tape];
                //psa/linar
                foreach (string item in dicz.Keys)
                {

                    DataTable dataTablez = dicz[item];
                    DataRow imageRow = dataTableImage.NewRow();
                    DataRow graphRow = dataTableImage.NewRow();
                    int i = 0;
                    for (; i < dataTablez.Rows.Count && i < 32; i++)
                    {
                        if (!tape.Equals("SAMPLE"))
                        {
                            var grimz = dataTablez.Rows[i]["Graph"];
                            if (grimz is Image)
                            {
                                graphRow[dataTableImage.Columns[i + 1]] = TDMK_ImageConverter.ImageToByteArray((Image)dataTablez.Rows[i]["Graph"], ImageFormat.Jpeg);
                            }
                            else
                            {
                                graphRow[dataTableImage.Columns[i + 1]] = dataTablez.Rows[i]["Graph"];

                            }
                        }
                        else
                        {
                            Image image = new Bitmap(20, 20);
                            graphRow[dataTableImage.Columns[i + 1]] = TDMK_ImageConverter.ImageToByteArray((Image)image, ImageFormat.Jpeg);

                        }
                        var imz = dataTablez.Rows[i]["Image"];
                        if (imz is Image)
                        {
                            imageRow[dataTableImage.Columns[i + 1]] = TDMK_ImageConverter.ImageToByteArray((Image)imz, ImageFormat.Jpeg);
                        }
                        else
                        {
                            imageRow[dataTableImage.Columns[i + 1]] = imz;
                        }
                    }
                    if (i < 32)
                    {
                        for (; i < 32; i++)
                        {
                            Image image = new Bitmap(20, 20);
                            if (tape.Equals("SAMPLE"))
                            {
                                imageRow[dataTableImage.Columns[i + 1]] = TDMK_ImageConverter.ImageToByteArray(image, ImageFormat.Jpeg);
                                graphRow[dataTableImage.Columns[i + 1]] = TDMK_ImageConverter.ImageToByteArray(image, ImageFormat.Jpeg);
                            }
                        }
                    }
                    imageRow["ID"] = pcs++;
                    if (!tape.Equals("SAMPLE"))
                    {
                        graphRow["ID"] = pcs;
                        pcs++;
                    }
                    dataTablez.Columns.Remove("Image");
                    dataTablez.Columns.Remove("Graph");
                    string json = ConverterService.DataTableToJson(dataTablez);
                    string key = $"&{tape}-{item}<ImageID@ {imageRow["ID"]}-{graphRow["ID"]}; Json@ {json}; Count@ {dataTablez.Rows.Count};>";
                    dataTableImage.Rows.Add(imageRow);
                    if (!tape.Equals("SAMPLE"))
                    {
                        dataTableImage.Rows.Add(graphRow);
                    }
                    resData += key;
                }

            }
            dr["Data"] = resData;
            dataTable.Rows.Add(dr);
            DataTable imageIDLIST = _dBContext.LoadDataTable(_TABLE_NAME, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo }, new[] { "ID", "Data" });

            string[] jsonZ = imageIDLIST.Rows[0]["Data"].ToString().Split('&');
            IList<string> idList = new List<string>();
            foreach (var item in jsonZ)
            {
                if (item.Contains("Before"))
                {
                    string content = item.Split('<')[1].TrimEnd('>').Trim().Split(':')[1].Trim();
                    idList.Add(content);
                }
                else
                {
                    string content = item.Split('<')[1].TrimEnd('>').Trim().Split(';')[0].Split('@')[1];
                    foreach (var item1 in content.Split('-'))
                    {
                        if (!string.IsNullOrEmpty(item1))
                        {
                            idList.Add(item1.Trim());
                        }
                    }
                }
            }
            res += _dBContext.BuckDataTable(dataTable, _TABLE_NAME, new[] { "ItemCode", "LotNo" }, null, "ID");
            res += _dBContext.DeleteData(_TABLE_NAME, "ID", idList.ToArray());
            res += _dBContext.SaveDataTable(dataTableImage, _TABLE_NAME + "_IMAGE", null, "ID");

            return res;
        }

        public Dictionary<string, Dictionary<string, DataTable>> Load(string itemCode, string lotNo, out DataTable before)
        {
            before = new DataTable();
            Dictionary<string, Dictionary<string, DataTable>> dic = new Dictionary<string, Dictionary<string, DataTable>>();
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            DataTable dt = _dBContext.LoadDataTable(_TABLE_NAME, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            if (dt.Rows.Count <= 0)
            {
                throw new Exception("Không có dữ liệu");
            }
            string json = dt.Rows[0]["Data"].ToString();
            string[] process = json.Split('&');
            foreach (var item in process)
            {
                if (item.Contains('<'))
                {

                    string pro = item.Split('<')[0];
                    switch (pro)
                    {
                        case "Before":
                            string proc = item.Split('<')[1];
                            string id = proc.Split(':')[1].TrimEnd('>').Trim();
                            DataTable image = _dBContext.LoadDataTable(_TABLE_NAME + "_IMAGE", new[] { "ID" }, new[] { id });
                            image.Columns.Remove("ID");
                            before = image;
                            break;

                        default:
                            string key = pro.Split('-')[0];
                            string tape = pro.Split('-')[1];
                            if (dic.TryGetValue(key, out Dictionary<string, DataTable> valueax))
                            {

                                if (!valueax.TryGetValue(tape, out DataTable data))
                                {
                                    valueax.Add(tape, MakeAgainDT(item.Split('<')[1].TrimEnd('>')));
                                }
                            }
                            else
                            {
                                Dictionary<string, DataTable> dicz = new Dictionary<string, DataTable>();
                                DataTable dz = MakeAgainDT(item.Split('<')[1].TrimEnd('>'));
                                dicz.Add(tape, dz);
                                dic.Add(key, dicz);
                            }
                            break;
                    }
                }
            }
            return dic;
        }

        private DataTable MakeAgainDT(string v)
        {
            DataTable imageDT = new DataTable();
            DataTable imageDT1 = new DataTable();
            DataTable result = new DataTable();
            int pcs = 0;
            string[] bson = v.Split(';');
            foreach (var item in bson)
            {
                string process = item.Split('@')[0].Trim();
                switch (process)
                {
                    case "Count":
                        pcs = int.Parse(item.Split('@')[1].Trim());
                        break;
                    case "ImageID":
                        string id = item.Split('@')[1].Trim();
                        string id1 = id.Split('-')[0];
                        string id2 = id.Split('-')[1];
                        imageDT = _dBContext.LoadDataTable(_TABLE_NAME + "_IMAGE", new[] { "ID" }, new[] { id1 });

                        if (!string.IsNullOrEmpty(id2))
                        {
                            imageDT1 = _dBContext.LoadDataTable(_TABLE_NAME + "_IMAGE", new[] { "ID" }, new[] { id2 });
                            imageDT1.Columns.Remove("ID");
                        }

                        imageDT.Columns.Remove("ID");
                        break;
                    case "Json":
                        string json = item.Split('@')[1].Trim();
                        result = ConverterService.JsonToDataTable(json);
                        break;
                }
            }
            result.Columns.Add("Image", typeof(byte[]));
            result.Columns.Add("Graph", typeof(byte[]));
            if (imageDT.Rows.Count <= 0)
            {
                return result;
            }
            List<DataRow> list = new List<DataRow>();
            for (int i = 0; i < pcs; i++)
            {
                DataRow row = result.Rows[i];
                if (i < 32)
                {
                    row["Image"] = (byte[])imageDT.Rows[0][result.Rows.IndexOf(row)];
                    if (imageDT1.Rows.Count > 0)
                    {
                        row["Graph"] = (byte[])imageDT1.Rows[0][result.Rows.IndexOf(row)];
                    }

                }
                else
                {
                    list.Add(row);
                }
            }
            if (list != null)
            {

                foreach (var item in list)
                {
                    result.Rows.Remove(item);
                }
            }
            return result;
        }
    }
}
