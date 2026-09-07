using OK2SHIP_SMT.Repositories;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using NationalInstruments.Restricted;
using OfficeOpenXml;

namespace OK2SHIP_SMT.Services
{
    public class XRayPictureService
    {
        private readonly string _NAME_SQL = "XRAY";
        private DBContext _dbContext = new DBContext();

        public static void GetInfor(string location, out string itemCode, out string lotNo)
        {
            string folderName = FileFolderRepository.GetFolderName(location);
            string[] splitFolderName = folderName.Split('-');
            itemCode = splitFolderName[2];
            lotNo = splitFolderName[3];
        }

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
                string _itemCode = lo[2];
                string _lotNo = lo[3] + (int.TryParse(lo[4], out int z) ? lo[4] : "");
                try
                {
                    _lotNo = int.Parse(_lotNo).ToString("D5");
                }
                catch
                {
                }

                if (!(itemCode.Equals(_itemCode) && _lotNo.Equals(lotNo)))
                {
                    return false;
                }
            }

            return true;
        }

        public int Save(DataTable dataTable, string itemCode, string lotNo, string maker, string type, int prime = -1)
        {
            if (dataTable.Rows.Count <= 0)
            {
                throw new Exception("Hãy get data!");
            }

            int res = 0;
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            maker = maker.Trim();
            int area = 0;
            if (prime == -1)
            {
                DataTable log = _dbContext.LoadDataTable(_NAME_SQL + "_NAS",
                    new[] { "ItemCode", "LotNo", "Maker", "Type" },
                    new[] { itemCode, lotNo, maker, type }, new[] { "ID", "Area" });
                if (log.Rows.Count > 0)
                {
                    string z = log.Rows[0]["Area"].ToString();
                    throw new Exception($" 1234 - Đã tồn tại dữ liệu bạn có muốn tiếp tục!");
                }
            }

            DataTable resDataTable = _dbContext.GetTableStructure(_NAME_SQL + "_NAS");
            NasRepository nas = new NasRepository();
            string location = nas.HandleImageDataTable(dataTable, _NAME_SQL, itemCode, lotNo);

            DataRow dr = resDataTable.NewRow();
            dr["ItemCode"] = itemCode;
            dr["LotNo"] = lotNo;
            dr["Maker"] = maker;
            dr["Data"] = ConverterService.DataTableToJson(dataTable);
            dr["Area"] = location;
            dr["Type"] = type;
            resDataTable.Rows.Add(dr);
            res += _dbContext.BuckDataTable(resDataTable, _NAME_SQL + "_NAS",
                new[] { "ItemCode", "LotNo", "Maker", "Type" },
                null, "ID");
            return res;
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

        public DataTable Read(string location, string itemCode, string lotNo, string type, bool prime = false)
        {
            type = type.Trim();
            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            if (!CheckNameFile(location, itemCode, lotNo))
            {
                throw new Exception("Vấn đề itemcode lotno");
            }

            string tableName = type.Trim().ToUpper().Replace(" ", "_");
            if (string.IsNullOrEmpty(tableName))
            {
                throw new Exception("Hãy chọn type");
            }

            DataTable pidList = _dbContext.LoadDataTable(tableName, new[] { "ItemCode", "LotNo" },
                new[] { itemCode, lotNo }, new[] { "ItemCode", "LotNo", "Net_No", "Pcs_No" });
            if (pidList.Rows.Count <= 0 && !prime)
            {
                throw new Exception($"1234 - {type} Chưa có productID!");
            }

            List<string> pid = new List<string>();
            foreach (DataRow ro in pidList.Rows)
            {
                if (ro["Net_No"].ToString().Trim().Equals("1"))
                {
                    pid.Add(ro["Pcs_No"].ToString());
                }
            }

            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("ID");
            dataTable.Columns.Add("Area");
            dataTable.Columns.Add("PT");
            dataTable.Columns.Add("Result");
            dataTable.Columns.Add("ProductID");
            dataTable.Columns.Add("Type");
            dataTable.Columns.Add("Image", typeof(byte[]));
            string[] listSubFolder = FileFolderRepository.GetSubFolders(location);
            foreach (string subFolderLocation in listSubFolder)
            {
                string[] listFile = FileFolderRepository.GetFileByExtension(subFolderLocation, "png");
                foreach (string fileLocation in listFile)
                {
                    try
                    {
                        string fileName = FileFolderRepository.GetFileName(fileLocation);
                        string[] spitFileName = fileName.Split('-');
                        DataRow row = dataTable.NewRow();
                        if (int.TryParse(spitFileName[2], out int number))
                        {
                            row["ID"] = spitFileName[2];
                            row["Result"] = spitFileName[3].Split('.')[0];
                        }
                        else
                        {
                            row["ID"] = "0";
                            row["Result"] = spitFileName[2].Split('.')[0];
                        }

                        row["PT"] = spitFileName[0];
                        row["Image"] = TDMK_ImageConverter.ImageFileToByteArray(fileLocation);
                        row["Area"] = spitFileName[1];
                        dataTable.Rows.Add(row);
                    }
                    catch
                    {
                        Debugger.Break();
                    }
                }
            }

            dataTable = dataTable.AsEnumerable().OrderBy(row =>
            {
                string ptValue = row.Field<string>("PT");
                // Lấy phần số sau chữ "PT" (bắt đầu từ ký tự thứ 2) và chuyển sang int
                if (int.TryParse(ptValue.Substring(2), out int number))
                {
                    return number;
                }

                return 0; // Giá trị mặc định nếu không parse được
            }).CopyToDataTable();
            return dataTable;
        }

        public DataTable Load(string itemCode, string lotNo, string maker, string type, bool prime = false)
        {
            if (string.IsNullOrEmpty(type.Trim()))
            {
                throw new Exception("Hãy chọn type!");
            }

            itemCode = itemCode.Trim();
            lotNo = lotNo.Trim();
            maker = maker.Trim();
            if (string.IsNullOrEmpty(itemCode) && string.IsNullOrEmpty(lotNo))
            {
                throw new Exception("Không được để trống itemcode lotno");
            }

            DataTable dataTable = _dbContext.LoadDataTable(_NAME_SQL + (prime ? "" : "_NAS"),
                new[] { "ItemCode", "LotNo", "Maker", "Type" }, new[] { itemCode, lotNo, maker, type }, null);
            if (dataTable.Rows.Count == 0)
            {
                throw new Exception("Không tìm thấy dữ liệu");
            }

            if (!prime)
            {
                DataTable res = dataTable;
                NasRepository _nas = new NasRepository();
                res = ConverterService.JsonToDataTable(dataTable.Rows[0]["Data"].ToString());
                _nas.MergeDataTable(res, _NAME_SQL, itemCode, lotNo, dataTable.Rows[0]["Area"].ToString());
                return res;
            }
            else
            {
                DataTable dataImage = _dbContext.LoadDataTable(_NAME_SQL + "_IMAGE", new[] { "Area" },
                    new[] { dataTable.Rows[0]["Area"].ToString() }, new[] { "Image", "ID" });
                DataTable dataTableRes =
                    ConvertDataTable(ConverterService.JsonToDataTable((string)dataTable.Rows[0]["Data"]), dataImage);
                return dataTableRes;
            }
        }

        public void ExportMultipleMakers(ExcelWorksheet workSheet, string itemCode, string lotNo, List<string> makers,
            string type, bool prime = false)
        {
            ExportProcess.CleanPhantomDimension(workSheet);
            // 1. Quét tìm template
            IDictionary<string, string> dicTemplate = ExportProcess.FindAddressByText(workSheet,
                new[] { "Bending 1", "FlexSN", "Sample", "Result" });

            string[] bendingAddresses = dicTemplate["Bending 1"].Split('-');
            string[] resultAddresses = dicTemplate["Result"].Split('-');

            int totalSlots = Math.Min(makers.Count, bendingAddresses.Length);

            // TỐI ƯU 1: CHỈ GỌI Dimension ĐÚNG 1 LẦN để giết "thủ phạm GetDimension" 129 giây
            int totalColumnsToCopy = 40; // Số cột an toàn mặc định
            if (workSheet.Dimension != null)
            {
                totalColumnsToCopy = Math.Min(workSheet.Dimension.End.Column, 40);
            }

            // Biến lưu trữ thông tin cho Pha 2
            var slotDataList = new List<dynamic>();

            // Biến cộng dồn dòng để giữ tọa độ chuẩn khi chạy Top-Down
            int rowOffset = 0;

            // ==========================================
            // PHA 1: XÂY DỰNG KHUNG (CHÈN DÒNG VÀ COPY STYLE)
            // ==========================================
            for (int slot = 0; slot < totalSlots; slot++) // DUYỆT TỪ TRÊN XUỐNG DƯỚI
            {
                string currentMaker = makers[slot];
                DataTable dt = Load(itemCode, lotNo, currentMaker, type, prime);

                var sortedAreaList = dt.AsEnumerable()
                    .Select(row => new { Area = row.Field<string>("Area"), ID = Convert.ToInt32(row["ID"]) })
                    .Distinct().OrderBy(x => x.Area).ThenBy(x => x.ID)
                    .Select(x => $"{x.Area}-{x.ID}").ToList();

                // Tính tọa độ thực tế của Template sau khi bị đẩy xuống bởi các Slot trước đó
                int originalBendingRow = workSheet.Cells[bendingAddresses[slot]].Start.Row;
                int originalResultRow = workSheet.Cells[resultAddresses[slot]].Start.Row;
                int startCol = workSheet.Cells[bendingAddresses[slot]].Start.Column;

                int currentBendingRow = originalBendingRow + rowOffset;
                int currentResultRow = originalResultRow + rowOffset;

                int n = sortedAreaList.Count - 1;

                if (n > 0)
                {
                    int startInsertRow = currentResultRow + 1;
                    int totalRowsToInsert = n * 2;

                    // LỆNH NÀY BÂY GIỜ SẼ SIÊU NHANH VÌ KHÔNG CÓ ẢNH ĐỂ ĐIỀU CHỈNH
                    workSheet.InsertRow(startInsertRow, totalRowsToInsert);

                    double heightBending = workSheet.Row(currentBendingRow).Height;
                    double heightResult = workSheet.Row(currentResultRow).Height;

                    var sourceBendingRange =
                        workSheet.Cells[currentBendingRow, 1, currentBendingRow, totalColumnsToCopy];
                    var sourceResultRange = workSheet.Cells[currentResultRow, 1, currentResultRow, totalColumnsToCopy];

                    for (int i = 0; i < n; i++)
                    {
                        int currentInsertPos = startInsertRow + (i * 2);
                        sourceBendingRange.Copy(workSheet.Cells[currentInsertPos, 1]);
                        sourceResultRange.Copy(workSheet.Cells[currentInsertPos + 1, 1]);

                        workSheet.Row(currentInsertPos).Height = heightBending;
                        workSheet.Row(currentInsertPos + 1).Height = heightResult;
                    }

                    // Cập nhật độ lệch dòng cho Slot tiếp theo
                    rowOffset += totalRowsToInsert;
                }

                // Lưu toàn bộ dữ liệu sạch vào Cache để dành cho Pha 2
                slotDataList.Add(new
                {
                    Maker = currentMaker,
                    DataTable = dt,
                    SortedAreas = sortedAreaList,
                    ActualBendingRow = currentBendingRow,
                    StartColumn = startCol,
                    SlotIndex = slot
                });
            }

            // ==========================================
            // PHA 2: ĐỔ DỮ LIỆU & HÌNH ẢNH (KHÔNG CÒN LỆNH INSERT ROW)
            // ==========================================
            foreach (var slotData in slotDataList)
            {
                int currentRowForFill = slotData.ActualBendingRow;
                string templateValueBend = workSheet.Cells[slotData.ActualBendingRow, slotData.StartColumn].Text;
                var dicLocalRowMapping = new Dictionary<string, int>();

                // Điền Header bên trái và Map dòng
                foreach (string item in slotData.SortedAreas)
                {
                    string areaName = item.Split('-')[0].Replace("B", "");
                    workSheet.Cells[currentRowForFill, slotData.StartColumn].Value =
                        templateValueBend.Replace("1", areaName);

                    dicLocalRowMapping[item] = currentRowForFill;
                    currentRowForFill += 2;
                }

                // Fill Kết quả và Hình Ảnh bằng tọa độ int
                foreach (DataRow row in slotData.DataTable.Rows)
                {
                    try
                    {
                        string key = $"{row["Area"]}-{row["ID"]}";
                        if (dicLocalRowMapping.TryGetValue(key, out int targetRow))
                        {
                            if (int.TryParse(row["PT"]?.ToString().Replace("PT", ""), out int colOffset))
                            {
                                int targetCol = slotData.StartColumn + colOffset;

                                // Fill Text
                                workSheet.Cells[targetRow + 1, targetCol].Value = row["Result"]?.ToString();

                                // Fill Ảnh
                                if (row["Image"] is byte[] imgBytes && imgBytes.Length > 0)
                                {
                                    string uniqueImgName = $"S{slotData.SlotIndex}_{key}_{colOffset}_{slotData.Maker}";
                                    ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[targetRow, targetCol],
                                        imgBytes, uniqueImgName);
                                }
                            }
                        }
                    }
                    catch (Exception exZ)
                    {
#if DEBUG
                        System.Diagnostics.Debug.WriteLine($"Lỗi: {exZ.Message}");
#endif
                    }
                }
            }
        }

        public string ExportMultipleMakers(string itemCode, string lotNo, List<string> makers, string type,
            bool prime = false)
        {
            ExportProcess exportProcess = new ExportProcess();
            string nameSheet = "X-Ray picture";
            switch (type)
            {
                case "Flex bending":
                    nameSheet = "Flex Bending - X-Ray pictures";
                    break;
                case "Thermal Cycling And Bend":
                    nameSheet = "TC & bending - X-Ray pictures";
                    break;
                case "Heat Soak And Bend":
                    nameSheet = "HS & bending - X-Ray pictures";
                    break;
            }

            using (ExcelPackage ex = exportProcess.FindFormatProcess(nameSheet, itemCode, lotNo))
            {
                using (ExcelWorksheet workSheet = exportProcess.FindSheet(ex, nameSheet))
                {
                    // 1. Quét tìm tất cả các vị trí Slot có sẵn trong Template
                    IDictionary<string, string> dicTemplate = ExportProcess.FindAddressByText(workSheet,
                        new[] { "Bending 1", "FlexSN", "Sample", "Result" });

                    // Cắt chuỗi để lấy mảng tọa độ của các slot (VD: ["A10", "A30"])
                    string[] bendingAddresses = dicTemplate["Bending 1"].Split('-');
                    string[] resultAddresses = dicTemplate["Result"].Split('-');

                    // Số slot thực tế xử lý = Số maker truyền vào (hoặc tối đa bằng số slot template hỗ trợ)
                    int totalSlots = Math.Min(makers.Count, bendingAddresses.Length);

                    // 2. DUYỆT TỪ DƯỚI LÊN TRÊN (VÒNG LẶP NGƯỢC) ĐỂ TRÁNH LỆCH DÒNG
                    for (int slot = totalSlots - 1; slot >= 0; slot--)
                    {
                        string currentMaker = makers[slot];

                        // Khởi tạo một Dictionary nội bộ cho riêng slot này để lưu tọa độ Fill dữ liệu
                        Dictionary<string, string> dicLocal = new Dictionary<string, string>();

                        string currentBendingAddr = bendingAddresses[slot];
                        string currentResultAddr = resultAddresses[slot];

                        // Load dữ liệu theo Maker hiện tại của Slot
                        DataTable dt = Load(itemCode, lotNo, currentMaker, type, prime);

                        // Sắp xếp và lọc dữ liệu (như code cũ)
                        List<string> sortedAreaList = dt.AsEnumerable()
                            .Select(row => new
                            {
                                Area = row.Field<string>("Area"),
                                ID = Convert.ToInt32(row["ID"])
                            })
                            .Distinct()
                            .OrderBy(x => x.Area)
                            .ThenBy(x => x.ID)
                            .Select(x => $"{x.Area}-{x.ID}")
                            .ToList();

                        int rowBending = workSheet.Cells[currentBendingAddr].Start.Row;
                        int rowResult = workSheet.Cells[currentResultAddr].Start.Row;
                        int totalColumns = workSheet.Dimension?.End.Column ?? 22;

                        int n = sortedAreaList.Count - 1;

                        // 3. TIẾN HÀNH CHÈN DÒNG VÀ COPY STYLE
                        if (n > 0)
                        {
                            int startInsertRow = rowResult + 1;
                            int totalRowsToInsert = n * 2;

                            // Chèn cục bộ (Bulk Insert)
                            workSheet.InsertRow(startInsertRow, totalRowsToInsert);

                            double heightBending = workSheet.Row(rowBending).Height;
                            double heightResult = workSheet.Row(rowResult).Height;

                            var sourceBendingRange = workSheet.Cells[rowBending, 1, rowBending, totalColumns];
                            var sourceResultRange = workSheet.Cells[rowResult, 1, rowResult, totalColumns];

                            for (int i = 0; i < n; i++)
                            {
                                int currentInsertPos = startInsertRow + (i * 2);
                                sourceBendingRange.Copy(workSheet.Cells[currentInsertPos, 1]);
                                sourceResultRange.Copy(workSheet.Cells[currentInsertPos + 1, 1]);

                                workSheet.Row(currentInsertPos).Height = heightBending;
                                workSheet.Row(currentInsertPos + 1).Height = heightResult;
                            }
                        }

                        // 4. GẮN GIÁ TRỊ BENDING (Cột tiêu đề)
                        string addressBendingForFill = currentBendingAddr;
                        string templateValueBend = workSheet.Cells[addressBendingForFill].Text;

                        foreach (string item in sortedAreaList)
                        {
                            string areaName = item.Split('-')[0].Replace("B", "");
                            workSheet.Cells[addressBendingForFill].Value = templateValueBend.Replace("1", areaName);

                            // Lưu tọa độ vào Dictionary nội bộ của slot này
                            dicLocal[item] = addressBendingForFill;
                            addressBendingForFill = ExportProcess.AddRow(addressBendingForFill, 2);
                        }

                        // 5. FILL DỮ LIỆU & HÌNH ẢNH
                        foreach (DataRow row in dt.Rows)
                        {
                            try
                            {
                                string key = $"{row["Area"]}-{row["ID"]}";
                                if (dicLocal.TryGetValue(key, out string cellMapped))
                                {
                                    if (int.TryParse(row["PT"]?.ToString().Replace("PT", ""), out int col))
                                    {
                                        string addressCell = ExportProcess.AddColumn(cellMapped, col);

                                        // Gán Result
                                        workSheet.Cells[ExportProcess.AddRow(addressCell, 1)].Value =
                                            row["Result"]?.ToString();

                                        // Gán Image (Kết hợp hàm Resize đã bàn trước đó nếu có)
                                        if (row["Image"] is byte[] imgBytes && imgBytes.Length > 0)
                                        {
                                            // Đặt tên ảnh kèm theo Slot để EPPlus không bị báo lỗi trùng tên object
                                            string uniqueImgName = $"S{slot}_{key}_{col}_{currentMaker}";
                                            ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressCell],
                                                imgBytes, uniqueImgName);
                                        }
                                    }
                                }
                            }
                            catch (Exception exZ)
                            {
#if DEBUG
                                System.Diagnostics.Debug.WriteLine($"Lỗi fill ảnh tại Slot {slot}: {exZ.Message}");
#endif
                            }
                        }
                    } // Kết thúc vòng lặp Slot

                    // 6. Lưu file sau khi tất cả các Slot đã được xử lý
                    string fileName = $"{itemCode.Trim()}-{lotNo.Trim()}_MultiMaker";
                    exportProcess.SaveExcelWorksheet(ex, nameSheet, fileName);
                }
            }

            return "Export báo cáo thành công!";
        }

        public void Export(ExcelWorksheet workSheet, string itemCode, string lotNo, string maker, string type,
            bool prime, int slot)
        {
            IDictionary<string, string> dic = ExportProcess.FindAddressByText(workSheet,
                new[] { "Bending 1", "FlexSN", "Sample", "Result" });
            dic["Bending 1"] = dic["Bending 1"].Split('-')[slot];
            dic["Result"] = dic["Result"].Split('-')[slot];
            DataTable dt = Load(itemCode, lotNo, maker, type, prime);
            List<string> sortedAreaList = dt.AsEnumerable()
                .Select(row => new
                {
                    Area = row.Field<string>("Area"),
                    // Chuyển ID sang int để sắp xếp đúng thứ tự số học (0, 1, 2...)
                    ID = Convert.ToInt32(row["ID"])
                })
                .Distinct() // Lọc các cặp Area-ID duy nhất
                .OrderBy(x => x.Area) // Sắp xếp Area tăng dần (A -> Z)
                .ThenBy(x => x.ID) // Sau đó sắp xếp ID tăng dần (0, 1, 2...)
                .Select(x => $"{x.Area}-{x.ID}") // Trả về định dạng Area-ID
                .ToList();

            int rowBending = workSheet.Cells[dic["Bending 1"]].Start.Row;
            int rowResult = workSheet.Cells[dic["Result"]].Start.Row;
            int totalColumns = workSheet.Dimension?.End.Column ?? 22;

            int n = sortedAreaList.Count - 1;
            int startInsertRow = rowResult + 1;
            double heightBending = workSheet.Row(rowBending).Height;
            double heightResult = workSheet.Row(rowResult).Height;
            for (int i = 0; i < n; i++)
            {
                // Tính toán vị trí chèn hiện tại
                int currentInsertPos = startInsertRow + (i * 2);

                // 1. Chèn 2 dòng trống
                workSheet.InsertRow(currentInsertPos, 2);

                // 2. Copy dữ liệu, định dạng (nội dung, style, công thức)
                workSheet.Cells[rowBending, 1, rowBending, totalColumns]
                    .Copy(workSheet.Cells[currentInsertPos, 1]);
                workSheet.Cells[rowResult, 1, rowResult, totalColumns]
                    .Copy(workSheet.Cells[currentInsertPos + 1, 1]);

                // 3. Gán lại chiều cao cho các hàng mới
                workSheet.Row(currentInsertPos).Height = heightBending;
                workSheet.Row(currentInsertPos + 1).Height = heightResult;
            }

            // gắn giá trị bending
            string addressBending = workSheet.Cells[dic["Bending 1"]].Address;
            foreach (string item in sortedAreaList)
            {
                string nameBend = item.Replace("B", "");
                string valueBend = workSheet.Cells[addressBending].Text;
                workSheet.Cells[addressBending].Value = valueBend.Replace("1", nameBend.Split('-')[0]);
                dic.Add(item, addressBending);
                addressBending = ExportProcess.AddRow(addressBending, 2);
            }

            // fill dữ liệu
            foreach (DataRow row in dt.Rows)
            {
                try
                {
                    string key = $"{row["Area"]}-{row["ID"]}";
                    if (dic.TryGetValue(key, out string value))
                    {
                        int col = int.Parse(row["PT"].ToString().Replace("PT", ""));
                        string addressCell = ExportProcess.AddColumn(value, col);
                        workSheet.Cells[ExportProcess.AddRow(addressCell, 1)].Value = row["Result"].ToString();
                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressCell], (byte[])row["Image"],
                            $"{key}-{col}");
                    }
                }
                catch (Exception exZ)
                {
                    Debugger.Break();
                }
            }
        }

        public string Export(string itemCode, string lotNo, string maker, string type, bool prime = false)
        {
            ExportProcess exportProcess = new ExportProcess();
            string nameSheet = "X-Ray picture";
            switch (type)
            {
                case "Flex bending":
                    nameSheet = "Flex Bending - X-Ray pictures";
                    break;
                case "Thermal Cycling And Bend":
                    nameSheet = "TC & bending - X-Ray pictures";
                    break;
                case "Heat Soak And Bend":
                    nameSheet = "HS & bending - X-Ray pictures";
                    break;
            }

            using (ExcelPackage ex = exportProcess.FindFormatProcess(nameSheet, itemCode, lotNo))
            {
                using (ExcelWorksheet workSheet = exportProcess.FindSheet(ex, nameSheet))
                {
                    IDictionary<string, string> dic = ExportProcess.FindAddressByText(workSheet,
                        new[] { "Bending 1", "FlexSN", "Sample", "Result" });
                    DataTable dt = Load(itemCode, lotNo, maker, type, prime);
                    List<string> sortedAreaList = dt.AsEnumerable()
                        .Select(row => new
                        {
                            Area = row.Field<string>("Area"),
                            // Chuyển ID sang int để sắp xếp đúng thứ tự số học (0, 1, 2...)
                            ID = Convert.ToInt32(row["ID"])
                        })
                        .Distinct() // Lọc các cặp Area-ID duy nhất
                        .OrderBy(x => x.Area) // Sắp xếp Area tăng dần (A -> Z)
                        .ThenBy(x => x.ID) // Sau đó sắp xếp ID tăng dần (0, 1, 2...)
                        .Select(x => $"{x.Area}-{x.ID}") // Trả về định dạng Area-ID
                        .ToList();

                    int rowBending = workSheet.Cells[dic["Bending 1"]].Start.Row;
                    int rowResult = workSheet.Cells[dic["Result"]].Start.Row;
                    int totalColumns = workSheet.Dimension?.End.Column ?? 22;

                    int n = sortedAreaList.Count - 1;
                    int startInsertRow = rowResult + 1;
                    double heightBending = workSheet.Row(rowBending).Height;
                    double heightResult = workSheet.Row(rowResult).Height;
                    for (int i = 0; i < n; i++)
                    {
                        // Tính toán vị trí chèn hiện tại
                        int currentInsertPos = startInsertRow + (i * 2);

                        // 1. Chèn 2 dòng trống
                        workSheet.InsertRow(currentInsertPos, 2);

                        // 2. Copy dữ liệu, định dạng (nội dung, style, công thức)
                        workSheet.Cells[rowBending, 1, rowBending, totalColumns]
                            .Copy(workSheet.Cells[currentInsertPos, 1]);
                        workSheet.Cells[rowResult, 1, rowResult, totalColumns]
                            .Copy(workSheet.Cells[currentInsertPos + 1, 1]);

                        // 3. Gán lại chiều cao cho các hàng mới
                        workSheet.Row(currentInsertPos).Height = heightBending;
                        workSheet.Row(currentInsertPos + 1).Height = heightResult;
                    }

                    // gắn giá trị bending
                    string addressBending = workSheet.Cells[dic["Bending 1"]].Address;
                    foreach (string item in sortedAreaList)
                    {
                        string nameBend = item.Replace("B", "");
                        string valueBend = workSheet.Cells[addressBending].Text;
                        workSheet.Cells[addressBending].Value = valueBend.Replace("1", nameBend.Split('-')[0]);
                        dic.Add(item, addressBending);
                        addressBending = ExportProcess.AddRow(addressBending, 2);
                    }

                    // fill dữ liệu
                    foreach (DataRow row in dt.Rows)
                    {
                        try
                        {
                            string key = $"{row["Area"]}-{row["ID"]}";
                            if (dic.TryGetValue(key, out string value))
                            {
                                if (int.TryParse(row["PT"]?.ToString().Replace("PT", ""), out int col))
                                {
                                    string addressCell = ExportProcess.AddColumn(value, col);

                                    // 1. Gán giá trị kết quả (Rất nhanh)
                                    workSheet.Cells[ExportProcess.AddRow(addressCell, 1)].Value =
                                        row["Result"]?.ToString();

                                    // 2. Xử lý ảnh: Lấy byte gốc -> Resize nhỏ gọn lại -> Đưa vào Excel
                                    if (row["Image"] is byte[] rawImgBytes && rawImgBytes.Length > 0)
                                    {
                                        // Truyền 0.7f để giữ lại 70% kích thước ảnh gốc (thu nhỏ 30%)
                                        byte[] optimizedImgBytes = ResizeImageBytes(rawImgBytes, 0.7f);

                                        ExportProcess.InsertImageToCell(workSheet, workSheet.Cells[addressCell],
                                            optimizedImgBytes, $"{key}-{col}");
                                    }
                                }
                            }
                        }
                        catch (Exception exZ)
                        {
#if DEBUG
                            Debugger.Break();
#endif
                        }
                    }


                    exportProcess.SaveExcelWorksheet(ex, nameSheet, $"{itemCode.Trim()}-{lotNo.Trim()}");
                }
            }

            return "Export thành công!";
        }

        private byte[] ResizeImageBytes(byte[] originalBytes, float scaleFactor = 0.7f)
        {
            if (originalBytes == null || originalBytes.Length == 0) return originalBytes;

            try
            {
                using (var ms = new MemoryStream(originalBytes))
                {
                    using (var img = Image.FromStream(ms))
                    {
                        // Tính toán kích thước mới dựa trên % tỷ lệ truyền vào
                        // Nếu thu nhỏ 30% -> giữ lại 70% -> scaleFactor = 0.7f
                        int targetWidth = (int)(img.Width * scaleFactor);
                        int targetHeight = (int)(img.Height * scaleFactor);

                        // Đảm bảo kích thước không bị tụt xuống 0
                        if (targetWidth <= 0) targetWidth = 1;
                        if (targetHeight <= 0) targetHeight = 1;

                        using (var bmp = new Bitmap(targetWidth, targetHeight))
                        {
                            using (var g = Graphics.FromImage(bmp))
                            {
                                // Cấu hình để vẽ lại cực nhanh
                                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.Low;
                                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighSpeed;
                                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighSpeed;

                                g.DrawImage(img, 0, 0, targetWidth, targetHeight);
                            }

                            using (var outMs = new MemoryStream())
                            {
                                // Giữ chất lượng nét ở mức 75%
                                var encoderParams = new EncoderParameters(1);
                                encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 75L);
                                var jpegCodec = ImageCodecInfo.GetImageEncoders()
                                    .FirstOrDefault(c => c.FormatID == ImageFormat.Jpeg.Guid);

                                if (jpegCodec != null)
                                {
                                    bmp.Save(outMs, jpegCodec, encoderParams);
                                }
                                else
                                {
                                    bmp.Save(outMs, ImageFormat.Jpeg);
                                }

                                return outMs.ToArray();
                            }
                        }
                    }
                }
            }
            catch
            {
                // Nếu lỗi resize, trả về ảnh gốc
                return originalBytes;
            }
        }
    }
}