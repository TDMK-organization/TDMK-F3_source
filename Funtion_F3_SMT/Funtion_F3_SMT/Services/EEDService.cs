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


namespace OK2SHIP_SMT.Services
{
    public class EEDService
    {
        private ExportProcess _export = new ExportProcess();
        private DBContext _dBContext = new DBContext();
        private const string FORMAT_NAME = "Environment en-durance";
        private const string _TABLE_NAME = "ENVIRONMENT_EN_DURANCE";

        public class InputModelExport
        {
            public string itemcode { get; set; }
            public string lotno { get; set; }
            public string maker { get; set; }

            public InputModelExport(string itemcode, string lotno, string maker)
            {
                this.itemcode = itemcode;
                this.lotno = lotno;
                this.maker = maker;
            }
        }

        public string Export(List<InputModelExport> listModel)
        {
            ExportProcess exportProcess = new ExportProcess();
            using (ExcelPackage ex = exportProcess.FindFormatProcess(FORMAT_NAME, listModel[0].itemcode, ""))
            {
                using (ExcelWorksheet workSheet = exportProcess.FindSheet(ex, FORMAT_NAME))
                {
                    bool firstTime = true;
                    // Xử lý từng model 
                    IDictionary<string, string> dicAddressTables = new Dictionary<string, string>();
                    foreach (InputModelExport inputModelExport in listModel)
                    {
                        // lấy dữ liệu của itemcode, lotno, maker tương ứng
                        DataTable dt = new DataTable();
                        Dictionary<string, Dictionary<string, DataTable>> dic = Load(inputModelExport.itemcode,
                            inputModelExport.lotno,
                            ref dt);

                        // Lấy danh sách tape của bộ dữ liệu
                        string[] listTape = dt.AsEnumerable()
                            .Where(row => row.Field<string>("Name") == "PSA") // Chữ 'Name' viết hoa chữ N
                            .Select(row => row.Field<string>("TAPE"))
                            .Distinct() // Lọc trùng lặp
                            .ToArray();

                        // Với lần đầu, setup format
                        if (firstTime)
                        {
                            IDictionary<string, string> address = ExportProcess.FindAddressByText(workSheet, new[]
                            {
                                "Liner peeling after ORT test",
                                "PSA peeling after ORT test",
                                "Sample",
                                "Summary Liner peeling force",
                                "Summary PSA peeling force",
                                "Liner_Maker_TAPE",
                                "PSA_Maker_TAPE"
                            });

                            int tablesToClone = listTape.Length - 1;

                            // Xác định cột bắt đầu và kết thúc của 1 bảng (VD: Từ cột A đến F -> 1 đến 6)
                            int startCol = workSheet.Cells[address["Liner peeling after ORT test"]].Start.Column;
                            int endCol = workSheet
                                .Cells[address["Sample"].Split('-')[address["Sample"].Split('-').Length - 1]].Start
                                .Column;
                            int blockWidth = endCol - startCol + 1;
                            int spaceCol = 2; // Khoảng cách giữa các bảng ngang
                            int rowSpacing = 2; // Khoảng cách giữa các maker dọc

                            // ================== TỌA ĐỘ LINER ==================
                            int linerTitleLargeRow = workSheet.Cells[address["Liner peeling after ORT test"]].Start.Row;
                            int linerStartRow = linerTitleLargeRow + 1;
                            int linerEndRow = workSheet.Cells[address["Summary Liner peeling force"]].Start.Row + 4;
                            int linerTitleRow = workSheet.Cells[address["Liner_Maker_TAPE"]].Start.Row;
                            int linerHeight = linerEndRow - linerStartRow + 1;

                            // ================== TỌA ĐỘ PSA ==================
                            int psaTitleLargeRow = workSheet.Cells[address["PSA peeling after ORT test"]].Start.Row;
                            int psaStartRow = psaTitleLargeRow + 1;
                            int psaEndRow = workSheet.Cells[address["Summary PSA peeling force"]].Start.Row + 4;
                            int psaTitleRow = workSheet.Cells[address["PSA_Maker_TAPE"]].Start.Row;
                            int psaHeight = psaEndRow - psaStartRow + 1;

                            int maxEndCol = endCol;

                            // ==========================================
                            // BƯỚC 1: DÀN HÀNG NGANG (THEO SỐ LƯỢNG TAPE)
                            // ==========================================
                            if (tablesToClone > 0)
                            {
                                for (int i = 1; i <= tablesToClone; i++)
                                {
                                    int destCol = startCol + i * (blockWidth + spaceCol);
                                    maxEndCol = destCol + blockWidth - 1;

                                    workSheet.Cells[linerTitleLargeRow, startCol, linerEndRow, endCol]
                                        .Copy(workSheet.Cells[linerTitleLargeRow, destCol]);
                                    workSheet.Cells[psaTitleLargeRow, startCol, psaEndRow, endCol]
                                        .Copy(workSheet.Cells[psaTitleLargeRow, destCol]);

                                    workSheet.Cells[linerTitleRow, destCol].Value =
                                        $"Liner_{listModel[0].maker}_{listTape[i]}";
                                    workSheet.Cells[psaTitleRow, destCol].Value =
                                        $"PSA_{listModel[0].maker}_{listTape[i]}";

                                    for (int c = 0; c < blockWidth; c++)
                                    {
                                        workSheet.Column(destCol + c).Width = workSheet.Column(startCol + c).Width;
                                    }
                                }
                            }

                            // Đổi tên cho bảng gốc (Maker đầu tiên) sau khi copy ngang xong
                            if (listTape.Length > 0)
                            {
                                workSheet.Cells[linerTitleRow, startCol].Value =
                                    $"Liner_{listModel[0].maker}_{listTape[0]}";
                                workSheet.Cells[psaTitleRow, startCol].Value =
                                    $"PSA_{listModel[0].maker}_{listTape[0]}";
                            }

                            // ==========================================
                            // BƯỚC 2: DÀN HÀNG DỌC (KIẾN TRÚC TỪ DƯỚI LÊN TRÊN AN TOÀN)
                            // ==========================================
                            int makersToClone = listModel.Count - 1;

                            if (makersToClone > 0)
                            {
                                // Để InsertRow an toàn, ta phải tính trước điểm thả (destRow) cho toàn bộ
                                // Ta sẽ tạo từ Maker cuối cùng (m = makersToClone) ngược về Maker 1
                                for (int m = makersToClone; m >= 1; m--)
                                {
                                    // === 2.1 COPY KHỐI PSA CỦA MAKER 'm' ===
                                    // Vị trí chèn = Đáy của PSA gốc + 1 + (m-1)*(chiều cao + khoảng trống) + khoảng trống
                                    int destRowPsa = psaEndRow + 1 + (m - 1) * (psaHeight + rowSpacing) + rowSpacing;

                                    workSheet.InsertRow(destRowPsa, psaHeight + rowSpacing);
                                    workSheet.Cells[psaStartRow, startCol, psaEndRow, maxEndCol]
                                        .Copy(workSheet.Cells[destRowPsa, startCol]);

                                    for (int r = 0; r < psaHeight; r++)
                                    {
                                        workSheet.Row(destRowPsa + r).Height = workSheet.Row(psaStartRow + r).Height;
                                    }

                                    for (int t = 0; t < listTape.Length; t++)
                                    {
                                        int c = startCol + t * (blockWidth + spaceCol);
                                        workSheet.Cells[destRowPsa + (psaTitleRow - psaStartRow), c].Value =
                                            $"PSA_{listModel[m].maker}_{listTape[t]}";
                                    }
                                }

                                // === 2.2 COPY KHỐI LINER CỦA MAKER 'm' ===
                                // Việc lặp chèn khối LINER phải diễn ra SAU khi toàn bộ PSA đã chèn xong.
                                // NHƯNG cũng phải lặp ngược từ dưới lên (m = makersToClone về 1) để không làm lệch LINER.
                                // Khi chèn LINER, nó sẽ đẩy toàn bộ cụm PSA vừa tạo xuống một cách hoàn hảo.
                                for (int m = makersToClone; m >= 1; m--)
                                {
                                    int destRowLiner = linerEndRow + 1 + (m - 1) * (linerHeight + rowSpacing) +
                                                       rowSpacing;

                                    workSheet.InsertRow(destRowLiner, linerHeight + rowSpacing);
                                    workSheet.Cells[linerStartRow, startCol, linerEndRow, maxEndCol]
                                        .Copy(workSheet.Cells[destRowLiner, startCol]);

                                    for (int r = 0; r < linerHeight; r++)
                                    {
                                        workSheet.Row(destRowLiner + r).Height =
                                            workSheet.Row(linerStartRow + r).Height;
                                    }

                                    for (int t = 0; t < listTape.Length; t++)
                                    {
                                        int c = startCol + t * (blockWidth + spaceCol);
                                        workSheet.Cells[destRowLiner + (linerTitleRow - linerStartRow), c].Value =
                                            $"Liner_{listModel[m].maker}_{listTape[t]}";
                                    }
                                }
                            }

                            // ==============================================================
                            // BƯỚC 3: XÓA DIMENSION ẢO CỘT BÊN PHẢI (Cực kỳ an toàn)
                            // ==============================================================
                            // Ta chỉ cắt bỏ cột dư, không cắt dòng dư để tránh rủi ro mất cụm PSA nằm ở dưới
                            if (workSheet.Dimension != null)
                            {
                                int currentMaxCol = workSheet.Dimension.End.Column;
                                if (currentMaxCol > maxEndCol)
                                {
                                    workSheet.DeleteColumn(maxEndCol + 1, currentMaxCol - maxEndCol);
                                }
                            }

                            // ==============================================================
                            // BƯỚC 4: TÌM KIẾM TỌA ĐỘ
                            // ==============================================================
                            List<string> listTitlesToFind = new List<string>();
                            foreach (var model in listModel)
                            {
                                foreach (var tape in listTape)
                                {
                                    listTitlesToFind.Add($"Liner_{model.maker}_{tape}".ToUpper());
                                    listTitlesToFind.Add($"PSA_{model.maker}_{tape}".ToUpper());
                                }
                            }

                            dicAddressTables = ExportProcess.FindAddressByText(workSheet, listTitlesToFind.ToArray());
                            firstTime = false;
                        }

                        foreach (string key in dic.Keys)
                        {
                            foreach (string keyZ in dic[key].Keys)
                            {
                                string addressDIC = $"{keyZ}_{inputModelExport.maker}_{key}".ToUpper();
                                if (dicAddressTables.TryGetValue(addressDIC, out string add))
                                {
                                    WriteData(workSheet, dic[key][keyZ], add);
                                }
                            }
                        }
                    }

                    exportProcess.SaveExcelWorksheet(ex, FORMAT_NAME, Guid.NewGuid().ToString());
                }
            }

            return "Export Successfully!";
        }

        public void WriteData(ExcelWorksheet worksheet, DataTable dataTable, string startAddress)
        {
            int i = 0;
            foreach (DataRow row in dataTable.Rows)
            {
                ExportProcess.InsertImageToCell(worksheet,
                    worksheet.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(startAddress, 2 + i), 1)],
                    (byte[])row["Image"],
                    Guid.NewGuid().ToString());
                ExportProcess.InsertImageToCell(worksheet,
                    worksheet.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(startAddress, 2 + i), 2)],
                    (byte[])row["Graph"],
                    Guid.NewGuid().ToString());
                try
                {
                    worksheet.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(startAddress, 2 + i), 3)].Value =
                        double.Parse(row["Peak(Gf)"].ToString());
                    worksheet.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(startAddress, 2 + i), 4)].Value =
                        double.Parse(row["Average(Gf)"].ToString());
                    worksheet.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(startAddress, 2 + i), 5)].Value =
                        double.Parse(row["Peak"].ToString());
                }
                catch
                {
                }

                worksheet.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(startAddress, 2 + i), 6)].Value =
                    double.Parse(row["Average"].ToString());
                worksheet.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(startAddress, 2 + i), 7)].Value =
                    row["Judgement Peeling force"].ToString();
                worksheet.Cells[ExportProcess.AddRow(ExportProcess.AddColumn(startAddress, 2 + i), 8)].Value =
                    row["Judgement failure mode"].ToString();
                i++;
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

            DataTable spec = _dBContext.LoadDataTable("SPEC_COMMENT_3", new[] { "ItemCode", "Sheet" },
                new[] { itemCode, "ENVIRONMENT_EN-DURANCE" });
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
                    string[] colHeader = new[]
                    {
                        "Liner peeling after ORT test", "PSA peeling after ORT test", sampleSTR, "Flex SN",
                        "Average Force"
                    };
                    IDictionary<string, string> dicHeader = ExportProcess.FindAddressByText(workSheet, colHeader);


                    Dictionary<string, string> dicCol = new Dictionary<string, string>();

                    #region Find area PSA Lineer

                    foreach (string item in new[] { "Liner peeling after ORT test", "PSA peeling after ORT test" })
                    {
                        dicCol.Add(item + "Flex SN", dicHeader[item]);
                        if (dicHeader.TryGetValue("Flex SN", out string value) &&
                            dicHeader.TryGetValue(item, out string valueZ))
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
                        if (dicHeader.TryGetValue("Average Force", out value) &&
                            dicHeader.TryGetValue(item, out valueZ))
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

                    string[] countPSA = dataTable.AsEnumerable().Where(row => row.Field<string>("name") == "PSA")
                        .Select(row => row.Field<string>("TAPE")).ToArray();
                    string[] countLiner = dataTable.AsEnumerable()
                        .Where(row => row.Field<string>("name") == "LINER")
                        .Select(row => row.Field<string>("TAPE")).ToArray();

                    if (dicCol.TryGetValue("PSA peeling after ORT testFlex SN", out string addressFlexSN)
                        && dicCol.TryGetValue($"PSA peeling after ORT test{sampleSTR}", out string addressSample)
                        && dicCol.TryGetValue("PSA peeling after ORT testAverage Force", out string addressCPK))
                    {
                        string addressPointer = workSheet.Cells[workSheet.Cells[addressCPK].Start.Row + 1,
                            workSheet.Cells[addressFlexSN].Start.Column].Address;

                        for (int i = 0; i < countPSA.Count() - 1; i++)
                        {
                            string point = addressPointer;
                            if (i == 0)
                            {
                                workSheet.Cells[ExportProcess.AddRow(addressFlexSN, 1)].Value =
                                    countPSA[countPSA.Count() - 1] + "%PSA";
                            }

                            string addressRange =
                                $"{addressFlexSN}:{workSheet.Cells[workSheet.Cells[addressCPK].Start.Row, workSheet.Cells[addressSample].Start.Column].Address}";
                            exportProcess.CopyAndInsert(workSheet, addressRange, ref addressPointer, true);
                            workSheet.Cells[ExportProcess.AddRow(point, 2)].Value = countPSA[i] + "%PSA";
                        }
                    }

                    if (dicCol.TryGetValue("Liner peeling after ORT testFlex SN", out addressFlexSN)
                        && dicCol.TryGetValue($"Liner peeling after ORT test{sampleSTR}", out addressSample)
                        && dicCol.TryGetValue("Liner peeling after ORT testAverage Force", out addressCPK))
                    {
                        string addressPointer = workSheet.Cells[workSheet.Cells[addressCPK].Start.Row + 1,
                            workSheet.Cells[addressFlexSN].Start.Column].Address;

                        for (int i = 0; i < countLiner.Count() - 1; i++)
                        {
                            string point = addressPointer;
                            if (i == 0)
                            {
                                workSheet.Cells[ExportProcess.AddRow(addressFlexSN, 1)].Value =
                                    countLiner[countLiner.Count() - 1] + "%LINER";
                            }

                            string addressRange =
                                $"{addressFlexSN}:{workSheet.Cells[workSheet.Cells[addressCPK].Start.Row, workSheet.Cells[addressSample].Start.Column].Address}";
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

                    IDictionary<string, string>
                        dicTape = ExportProcess.FindAddressByText(workSheet, tapeList.ToArray());

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

        private void FillDataInTape(ExcelWorksheet worksheet, DataTable dataTable, DataTable spec, string address,
            int sample, string name, string tape)
        {
            string addressImageSample =
                ExportProcess.getRangeBaseAddressByCellAddress(worksheet, ExportProcess.AddRow(address, 1));
            byte[] imgBck = (byte[])((DataRow)spec.AsEnumerable().FirstOrDefault(row =>
                row.Field<string>("Name") == name && row.Field<string>("TAPE") == tape))["Image Sample"];
            ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[addressImageSample], imgBck,
                $"BACK{tape}-{name}");
            int id = 0;
            string sampleAddress = ExportProcess.AddColumn(address, 2);
            foreach (DataRow dataRow in dataTable.Rows)
            {
                if (id >= sample)
                {
                    break;
                }

                string image = ExportProcess.AddRow(sampleAddress, 1);

                ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[image],
                    (byte[])dataRow["Image"],
                    $"Image{tape}-{name}-{id}");
                string graph = ExportProcess.AddRow(image, 1);
                ExportProcess.InsertImageToCell(worksheet, worksheet.Cells[graph],
                    (byte[])dataRow["Graph"],
                    $"Graph{tape}-{name}-{id}");
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

            return ValidateService.compareItemCodeLotNo(itemCodeL, $"{itemCode}-{lotNo}");
            ;
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
                    IList<KeyValuePair<Image, string>> listImage =
                        FileFolderRepository.ListAllPictureInAFolder(item + "\\ANH");
                    var orderList = listImage.OrderBy(x => int.Parse(x.Value.Split('.')[0]));

                    int excelIndex = 1; // Khởi tạo số thứ tự để map với file Excel (1, 2, 3...)

                    foreach (var itemImage in orderList)
                    {
                        DataRow row = dataTable.NewRow();

                        // Vẫn giữ lại số của ảnh (ví dụ 60, 61...) để lưu vào cột PCS nếu bạn muốn
                        string numPcs = itemImage.Value.Split('.')[0].Trim();
                        row["PCS"] = excelIndex;
                        row["Image"] = itemImage.Key;
                        row["Tape"] = tape;

                        // Xử lý ưu tiên lấy file dạng {excelIndex}-1.xlsx thay vì {excelIndex}.xlsx
                        string targetExcelPath = $"{item}\\{excelIndex}-1.xlsx";

                        // Nếu file x-1 không tồn tại thì fallback về file x.xlsx bình thường
                        if (!System.IO.File.Exists(targetExcelPath))
                        {
                            targetExcelPath = $"{item}\\{excelIndex}.xlsx";
                        }

                        // Chèn đường dẫn mới vào hàm
                        row["Graph"] = SolveFileLinerPSQ(targetExcelPath, out string resFile);

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

                        // Tăng số thứ tự Excel cho ảnh tiếp theo
                        excelIndex++;
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
                        IDictionary<string, string> dic =
                            ExportProcess.FindAddressByText(workSheet, new[] { "Max", "Average" }, true);
                        if (!dic.TryGetValue("Max", out string max))
                        {
                            max = "";
                        }

                        if (!dic.TryGetValue("Average", out string min))
                        {
                            min = "";
                        }

                        int i = 0;
                        while (true)
                        {
                            try
                            {
                                max = workSheet.Cells[ExportProcess.AddColumn(dic["Max"], i)].Text;
                                if (double.TryParse(max, out double _))
                                {
                                    break;
                                }

                                i++;
                            }
                            catch
                            {
                            }
                        }

                        i = 0;
                        while (true)
                        {
                            try
                            {
                                min = workSheet.Cells[ExportProcess.AddColumn(dic["Average"], i)].Text;

                                if (double.TryParse(min, out double _))
                                {
                                    break;
                                }

                                i++;
                            }
                            catch
                            {
                            }
                        }

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

        public void SolveFolderPSALiner(string location, Dictionary<string, Dictionary<string, DataTable>> dic,
            DataTable spec, string processz, int pcs = -1)
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

        public int save(string itemCode, string lotNo, DataTable before,
            Dictionary<string, Dictionary<string, DataTable>> dic, int prime = -1)
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
                DataTable checker = _dBContext.LoadDataTable(_TABLE_NAME, new[] { "ItemCode", "LotNo" },
                    new[] { itemCode, lotNo }, new[] { "ID", "Data" });
                if (checker.Rows.Count > 0)
                {
                    DataRow row = checker.Rows[0];
                    throw new Exception(
                        $"1234 - {row["Data"].ToString().Split('#')[0]} - ItemCode and LotNo already exist");
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

        public Dictionary<string, Dictionary<string, DataTable>> Load(string itemCode, string lotNo,
            ref DataTable before, bool legacy = false)
        {
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(lotNo))
            {
                throw new Exception("ItemCode or LotNo is null");
            }

            DataTable table =
                _dBContext.LoadDataTable(_TABLE_NAME, new[] { "ItemCode", "LotNo" }, new[] { itemCode, lotNo });
            if (table.Rows.Count <= 0)
            {
                throw new Exception("Không có dữ liệu của itemcode lotno");
            }

            string[] json = table.Rows[0]["Data"].ToString().Trim().Split('#');
            string area = json[0];
            DataTable tableImage = new DataTable();
            if (legacy)
            {
                tableImage = _dBContext.LoadDataTable(_TABLE_NAME + "_IMG", new[] { "Area" }, new[] { area },
                    new[] { "ID", "Image" });
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

            Dictionary<string, Dictionary<string, DataTable>> dic =
                new Dictionary<string, Dictionary<string, DataTable>>();
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

        public void FillProductID(Dictionary<string, Dictionary<string, DataTable>> dic, string itemCode, string lotNo,
            string location)
        {
            if (string.IsNullOrEmpty(location.Trim()))
            {
                throw new Exception("Điền đường dẫn productID");
            }

            ProductIDService productIDService = new ProductIDService(itemCode, lotNo, location,
                new[] { "ORT", "environment" }, new[] { "LINER", "PSA" });
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
            DataTable dataTable = _dBContext.LoadDataTable("SPEC_COMMENT_3", new[] { "ItemCode", "Sheet" },
                new[] { itemCode, "ENVIRONMENT_EN-DURANCE" });
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